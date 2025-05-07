import os
import zipfile
import shutil
import io
from googleapiclient.http import MediaIoBaseDownload
from StudentSubmission import StudentSubmission
from utils import extract_prefix, get_submission_timestamp, calculate_delay, log_error, get_due_date, log_info


def download_submissions(classroom_service, drive_service, submissions, download_folder,
                         classroom_id, coursework_id, questions_data, num_questions):
    students = []
    due_date = get_due_date(classroom_service, classroom_id, coursework_id)

    for submission in submissions.get('studentSubmissions', []):
        try:
            student_id = submission['userId']
            student = classroom_service.courses().students().get(courseId=classroom_id, userId=student_id).execute()
            student_email = student['profile']['emailAddress']
            student_login = extract_prefix(student_email)
            student_name = student['profile']['name']['fullName']

            entregou = 1
            atrasou = 0
            formatacao = 1
            comentarios = ''
            copia = 0

            submission_date = get_submission_timestamp(submission, student_id)
            attachments = submission.get('assignmentSubmission', {}).get('attachments', [])
            student_folder = os.path.join(download_folder, student_login)
            os.makedirs(student_folder, exist_ok=True)

            if not attachments:
                entregou = 0
                comentarios = "Não entregou a atividade."
            else:
                atrasou = calculate_delay(due_date, submission_date) if due_date and submission_date else 0

                for attachment in attachments:
                    file_id = attachment.get('driveFile', {}).get('id')
                    file_name = attachment.get('driveFile', {}).get('title')
                    request = drive_service.files().get_media(fileId=file_id)
                    file_path = os.path.join(student_folder, file_name)

                    with io.FileIO(file_path, 'wb') as fh:
                        downloader = MediaIoBaseDownload(fh, request)
                        done = False
                        while not done:
                            status, done = downloader.next_chunk()

                    if file_name.endswith('.rar'):
                        formatacao = 0
                        comentarios += f" Enviou .rar ({file_name}) ao invés de .zip."
                    elif file_name.endswith('.zip') and file_name != f"{student_login}.zip":
                        formatacao = 0
                        comentarios += f" Nome do zip incorreto: {file_name}."

            student_obj = StudentSubmission(
                name=student_name,
                email=student_email,
                login=student_login,
                entregou=entregou,
                atrasou=atrasou,
                formatacao=formatacao,
                copia=copia,
                comentario=comentarios.strip()
            )
            students.append(student_obj)

        except Exception as e:
            log_error(f"Erro ao processar submissão de aluno: {e}")

    return students


def organize_extracted_files(base_folder):
    try:
        submissions_folder = os.path.join(base_folder, "submissions")
        os.makedirs(submissions_folder, exist_ok=True)

        for item in os.listdir(base_folder):
            item_path = os.path.join(base_folder, item)
            if item == "submissions" or item.endswith(".json"):
                continue

            if zipfile.is_zipfile(item_path):
                extraction_path = os.path.join(submissions_folder, item.replace('.zip', ''))
                os.makedirs(extraction_path, exist_ok=True)
                with zipfile.ZipFile(item_path, 'r') as zip_ref:
                    zip_ref.extractall(extraction_path)
                os.remove(item_path)
            elif os.path.isdir(item_path):
                shutil.move(item_path, os.path.join(submissions_folder, item))
    except Exception as e:
        log_error(f"Erro ao organizar arquivos extraídos: {e}")


def move_non_zip_files(download_folder):
    try:
        submissions_folder = os.path.join(download_folder, 'submissions')
        for item in os.listdir(download_folder):
            item_path = os.path.join(download_folder, item)
            if os.path.isdir(item_path) and item != 'submissions':
                destination_folder = os.path.join(submissions_folder, item)
                if not os.path.exists(destination_folder):
                    os.rename(item_path, destination_folder)
    except Exception as e:
        log_error(f"Erro ao mover arquivos que não estavam zipados: {e}")


def if_there_is_a_folder_inside(submissions_folder):
    try:
        def move_files_to_initial_folder(first_folder, folder_name):
            if os.path.basename(first_folder).startswith('.'):
                return

            items = os.listdir(first_folder)
            subfolders = [item for item in items if os.path.isdir(os.path.join(first_folder, item)) and not item.startswith('.')]

            for subfolder in subfolders:
                subfolder_path = os.path.join(first_folder, subfolder)
                move_files_to_initial_folder(subfolder_path, folder_name)

            files = [item for item in items if os.path.isfile(os.path.join(first_folder, item)) and not item.startswith('.')]
            for file in files:
                file_path = os.path.join(first_folder, file)
                destination = os.path.join(submissions_folder, os.path.basename(first_folder), file)
                os.makedirs(os.path.dirname(destination), exist_ok=True)
                shutil.move(file_path, destination)

            if first_folder != submissions_folder and not os.listdir(first_folder):
                os.rmdir(first_folder)

        for folder_name in os.listdir(submissions_folder):
            folder_path = os.path.join(submissions_folder, folder_name)
            if os.path.isdir(folder_path) and not folder_name.startswith('.'):
                move_files_to_initial_folder(folder_path, folder_name)
    except Exception as e:
        log_error(f"Erro ao processar subpastas: {e}")


def delete_subfolders_in_student_folders(submissions_folder):
    try:
        for student_folder in os.listdir(submissions_folder):
            student_folder_path = os.path.join(submissions_folder, student_folder)
            if os.path.isdir(student_folder_path):
                for item in os.listdir(student_folder_path):
                    item_path = os.path.join(student_folder_path, item)
                    if os.path.isdir(item_path):
                        shutil.rmtree(item_path)
    except Exception as e:
        log_error(f"Erro ao deletar subpastas nas pastas dos estudantes: {e}")


def remove_empty_folders(folder):
    try:
        for dirpath, dirnames, filenames in os.walk(folder, topdown=False):
            if not dirnames and not filenames:
                os.rmdir(dirpath)
    except Exception as e:
        log_error(f"Erro ao remover pastas vazias: {e}")
