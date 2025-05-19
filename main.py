import os
import re
import shutil
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload
from StudentSubmission import StudentSubmission, save_students_to_txt, load_students_from_txt
from file_renamer import rename_files, integrate_renaming
from submission_handler import download_submissions
from utils import log_error, format_list_title, read_id_from_file
from google_auth_utils import get_credentials, get_gspread_client
from classroom_utils import list_classroom_data
from sheet_id_handler import  list_informations, list_questions
from ListMetadata import ListMetadata, save_metadata_to_json, load_metadata_from_json
from folders_organizer import (organize_extracted_files, move_non_zip_files, if_there_is_a_folder_inside, delete_subfolders_in_student_folders, remove_empty_folders)
from dataclasses import dataclass, asdict, is_dataclass

def main():
    try:
        creds = get_credentials()
        classroom_service = build("classroom", "v1", credentials=creds)
        drive_service = build("drive", "v3", credentials=creds)

        sheet_id = read_id_from_file("sheet_id.txt")
        if not sheet_id:
            print("Arquivo 'sheet_id.txt' não encontrado.\n")
            return

        semester, lists = list_informations(sheet_id)

        list_name = list_title = None
        list_title_a = None
        turma_folders = []

        for class_letter in ["A", "B"]:
            turma_type = f"TURMA {class_letter}"

            classroom_id, coursework_id, classroom_name, list_name, list_title = list_classroom_data(
                classroom_service, semester, turma_type=turma_type, saved_assignment_title=list_title_a
            )

            if not classroom_id:
                print("Dados da turma não encontrados.\n")
                return

            if class_letter == "A":
                list_title_a = list_title
                list_name_ref = list_name
                formatted_list = format_list_title(list_name)
            else:
                if list_name != list_name_ref:
                    print("As duas turmas devem usar a mesma atividade.\n")
                    return

            try:
                questions_data, num_questions, score = list_questions(sheet_id, list_name)
                if not score or not questions_data:
                    print("\nA aba da planilha precisa conter número de questões e score.")
                    return
            except Exception as e:
                print(f"Erro ao carregar dados da planilha: {e}")
                return

            formatted_class = f"turma{class_letter}"
            script_dir = os.path.dirname(os.path.abspath(__file__))
            base_path = os.path.join(script_dir, "Downloads", formatted_list)
            zips_folder = os.path.join(base_path, f"zips_{formatted_class}")
            submissions_folder = os.path.join(zips_folder, f"submissions_{formatted_class}")

            os.makedirs(zips_folder, exist_ok=True)

            # Salva apenas a pasta zips para o integrate_renaming
            turma_folders.append(zips_folder)

            metadata = ListMetadata(
                class_name=classroom_name,
                list_name=list_title,
                num_questions=num_questions,
                score=score
            )

            metadata_filename = f"metadata_turma{class_letter.upper()}.json"
            metadata_path = os.path.join(base_path, metadata_filename)
            save_metadata_to_json(metadata, metadata_path)

            submissions = classroom_service.courses().courseWork().studentSubmissions().list(
                courseId=classroom_id, courseWorkId=coursework_id).execute()

            student_list = download_submissions(classroom_service, drive_service, submissions, zips_folder, classroom_id, coursework_id)
            print("\n\nDownload completo. Arquivos salvos em:", os.path.abspath(zips_folder))

            students_filename = f"students_turma{class_letter.upper()}.json"
            students_path = os.path.join(base_path, students_filename)
            save_students_to_txt(student_list, students_path)

            organize_extracted_files(zips_folder, student_list, formatted_class)
            move_non_zip_files(zips_folder, formatted_class)
            if_there_is_a_folder_inside(student_list, submissions_folder)
            delete_subfolders_in_student_folders(submissions_folder)
            remove_empty_folders(submissions_folder)
            save_students_to_txt(student_list, students_path)

            print("\nProcesso de organização de pastas finalizado:", os.path.abspath(submissions_folder))

            rename_files(submissions_folder, list_title, questions_data, student_list)
            save_students_to_txt(student_list, students_path)
            print("\nProcesso de verificação e renomeação finalizado.")

        # Renomeia arquivos finais
        integrate_renaming(turma_folders, list_title, questions_data)

        # Junta tudo em uma única pasta de submissões
        final_submissions_folder = os.path.join(script_dir, "Downloads", formatted_list, "submissions")
        os.makedirs(final_submissions_folder, exist_ok=True)

        for zips_folder in turma_folders:
            class_name = os.path.basename(zips_folder).replace("zips_", "")
            src_submission_path = os.path.join(zips_folder, f"submissions_{class_name}")
            for student in os.listdir(src_submission_path):
                src_path = os.path.join(src_submission_path, student)
                dst_path = os.path.join(final_submissions_folder, student)
                shutil.move(src_path, dst_path)

        print("\nSubmissões unificadas em:", final_submissions_folder)

    except Exception as e:
        log_error(f"Erro no fluxo principal: {e}")

if __name__ == "__main__":
    main()