from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload
from StudentSubmission import StudentSubmission, save_students_to_txt, load_students_from_txt
from file_renamer import integrate_renaming
from submission_handler import download_submissions, organize_extracted_files, move_non_zip_files, if_there_is_a_folder_inside, delete_subfolders_in_student_folders, remove_empty_folders
from utils import log_error, format_list_title, read_id_from_file
from google_auth_utils import get_credentials, get_gspread_client
from spreadsheet_handler import (create_or_get_google_sheet_in_folder, header_worksheet, insert_header_title, freeze_and_sort, fill_worksheet_with_students)
from classroom_utils import list_classroom_data
from sheet_id_handler import  list_informations, list_questions
import re
import os
import shutil

def main():
    try:
        creds = get_credentials()
        classroom_service = build("classroom", "v1", credentials=creds)
        drive_service = build("drive", "v3", credentials=creds)

        sheet_id = read_id_from_file("sheet_id.txt")
        if not sheet_id:
            print("Arquivo 'sheet_id.txt' com o ID da planilha não encontrado. Não é possível rodar o script sem essa planilha.\n")
            return

        semester, lists = list_informations(sheet_id)

        all_students = []
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
                    print("\nA aba da planilha precisa estar preenchida com o número de questões e o score.")
                    return
            except Exception as e:
                print(f"Erro ao carregar dados da planilha: {e}")
                return

            formatted_class = f"turma{class_letter}"
            download_folder = os.path.join("Downloads", f"download_{formatted_class}_{formatted_list}")
            turma_folders.append(f"download_{formatted_class}_{formatted_list}")
            os.makedirs(download_folder, exist_ok=True)

            submissions = classroom_service.courses().courseWork().studentSubmissions().list(
                courseId=classroom_id, courseWorkId=coursework_id).execute()

            student_list = download_submissions(classroom_service, drive_service, submissions, download_folder, classroom_id, coursework_id)

            save_students_to_txt(student_list, os.path.join(download_folder, "students.json"))

            organize_extracted_files(download_folder)
            move_non_zip_files(download_folder)
            submissions_folder = os.path.join(download_folder, 'submissions')
            if_there_is_a_folder_inside(submissions_folder)
            delete_subfolders_in_student_folders(submissions_folder)
            remove_empty_folders(submissions_folder)

        # Integrar renomeação e salvar students_final.json
        integrate_renaming(turma_folders, list_title, questions_data)
        final_students_path = os.path.join("Downloads", "students_final.json")
        final_students = load_students_from_txt(final_students_path)

        # Unir submissões
        final_submissions_folder = os.path.join("Downloads", "submissions")
        os.makedirs(final_submissions_folder, exist_ok=True)

        for folder in turma_folders:
            src = os.path.join("Downloads", folder, "submissions")
            for student in os.listdir(src):
                src_path = os.path.join(src, student)
                dst_path = os.path.join(final_submissions_folder, student)
                shutil.move(src_path, dst_path)

        folder_id = read_id_from_file("folder_id.txt")
        if not folder_id:
            print("Arquivo 'folder_id.txt' não encontrado ou inválido.\n")
            return

        worksheet = create_or_get_google_sheet_in_folder("PIF", list_name, folder_id)
        header_worksheet(worksheet, num_questions, score)

        for student in final_students:
            worksheet.append_rows([student.to_list(num_questions)])

        freeze_and_sort(worksheet)
        insert_header_title(worksheet, "PIF", list_title)
        print("\nProcesso finalizado com sucesso.\n")

    except Exception as e:
        log_error(f"Erro no fluxo principal: {e}")

if __name__ == "__main__":
    main()
