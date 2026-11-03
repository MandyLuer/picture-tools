"""Read-only DOC conversion worker. Never attaches to the user's Word instance."""
import json
from pathlib import Path
import subprocess
import sys
import uuid


def convert_with_word(source, target, output_format=16):
    import pythoncom
    import win32com.client
    import win32api
    import win32job
    import win32process

    pythoncom.CoInitialize()
    application = document = None
    original_security = None
    job = None
    try:
        existing_processes = set(win32process.EnumProcesses())
        application = win32com.client.DispatchEx('Word.Application')
        # Word.Application has no Hwnd property. Locate only newly created Word
        # processes, never an already running instance belonging to the user.
        new_word_processes = []
        for process_id in set(win32process.EnumProcesses()) - existing_processes:
            handle = None
            try:
                handle = win32api.OpenProcess(0x410, False, process_id)
                executable = win32process.GetModuleFileNameEx(handle, 0)
                if Path(executable).name.casefold() == 'winword.exe':
                    new_word_processes.append(process_id)
            except Exception:
                continue
            finally:
                if handle is not None:
                    handle.Close()
        if len(new_word_processes) == 1:
            process_id = new_word_processes[0]
            job = win32job.CreateJobObject(None, '')
            limits = win32job.QueryInformationJobObject(job, win32job.JobObjectExtendedLimitInformation)
            limits['BasicLimitInformation']['LimitFlags'] |= win32job.JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
            win32job.SetInformationJobObject(job, win32job.JobObjectExtendedLimitInformation, limits)
            handle = win32api.OpenProcess(0x100 | 0x1, False, process_id)
            try:
                win32job.AssignProcessToJobObject(job, handle)
            finally:
                handle.Close()
        application.Visible = False
        application.DisplayAlerts = 0
        original_security = application.AutomationSecurity
        application.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
        password = uuid.uuid4().hex  # Fail encrypted documents instead of showing a password prompt.
        document = application.Documents.Open(FileName=str(source), ConfirmConversions=False,
                    ReadOnly=True, AddToRecentFiles=False, Visible=False, NoEncodingDialog=True,
                    PasswordDocument=password, PasswordTemplate=password, WritePasswordDocument=password)
        document.SaveAs2(FileName=str(target), FileFormat=output_format, AddToRecentFiles=False)
    except Exception as exc:
        raise ValueError('DOC 转换失败，请确认文档未损坏、未加密，且 Microsoft Word 已完成安装和激活') from exc
    finally:
        try:
            if document is not None:
                document.Close(SaveChanges=0)
        finally:
            try:
                if application is not None:
                    if original_security is not None:
                        application.AutomationSecurity = original_security
                    application.Quit(SaveChanges=0)
            finally:
                if job is not None:
                    job.Close()
                pythoncom.CoUninitialize()


def convert_with_libreoffice(executable, root):
    output = root / 'converted'
    output.mkdir()
    profile = (root / 'lo-profile').as_uri()
    completed = subprocess.run([executable, f'-env:UserInstallation={profile}', '--headless', '--norestore',
                               '--convert-to', 'docx', '--outdir', str(output), str(root / 'input.doc')],
                              timeout=75, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                              creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == 'win32' else 0)
    converted = output / 'input.docx'
    if completed.returncode or not converted.is_file():
        raise ValueError('LibreOffice 无法转换该 DOC，请检查文档是否损坏或加密')
    converted.replace(root / 'output.docx')


def main(request_file):
    root = Path(request_file).resolve().parent
    try:
        request = json.loads(Path(request_file).read_text(encoding='utf-8'))
        if request['provider'] == 'word':
            convert_with_word(root / 'input.doc', root / 'output.docx')
        else:
            convert_with_libreoffice(request['provider'], root)
        result, status = {'converted': True}, 0
    except Exception as exc:
        result, status = {'error':str(exc)}, 1
    (root / 'result.json').write_text(json.dumps(result, ensure_ascii=False), encoding='utf-8')
    return status


if __name__ == '__main__':
    sys.exit(main(sys.argv[1]))
