import os
import shutil
from utils import log_error, log_info
from StudentSubmission import save_students_to_txt, load_students_from_txt

def rename_files(submissions_folder, list_title, questions_data, students):
    try:
        if 'HASKELL' in list_title.upper():
            check_for_missing_files(submissions_folder, students, extension=".hs")
            rename_files_based_on_dictionary(submissions_folder, questions_data, students, extension=".hs")
            return 'haskell'
        else:
            check_for_missing_files(submissions_folder, students, extension=".c")
            rename_files_based_on_dictionary(submissions_folder, questions_data, students, extension=".c")
            return 'c'
    except Exception as e:
        log_error(f"Erro ao renomear arquivos: {e}")
        return None

def check_for_missing_files(folder, students, extension):
    try:
        for student in students:
            student_folder = os.path.join(folder, student.login)
            if os.path.exists(student_folder):
                files = [f for f in os.listdir(student_folder) if f.endswith(extension)]
                if not files:
                    student.entregou = 0
                    student.comentario += f" Não há arquivos {extension} na submissão."
    except Exception as e:
        log_error(f"Erro ao verificar arquivos {extension}: {e}")

def rename_files_based_on_dictionary(submissions_folder, questions_dict, students, extension):
    try:
        for student in students:
            student_folder = os.path.join(submissions_folder, student.login)
            if not os.path.exists(student_folder):
                continue

            for file_name in os.listdir(student_folder):
                if not file_name.endswith(extension):
                    continue

                original_path = os.path.join(student_folder, file_name)
                renamed = False

                for question_number, possible_names in questions_dict.items():
                    for possible_name in possible_names:
                        if possible_name.lower() in file_name.lower():
                            new_name = f"q{question_number}{extension}"
                            new_path = os.path.join(student_folder, new_name)
                            os.rename(original_path, new_path)
                            renamed = True
                            break
                    if renamed:
                        break

                if not renamed:
                    student.comentario += f" Arquivo {file_name} não corresponde a nenhuma questão."
                    log_info(f"Arquivo não renomeado para {student.login}: {file_name}")
    except Exception as e:
        log_error(f"Erro ao renomear arquivos com base no dicionário: {e}")

def integrate_renaming(turmas, list_title, questions_data):
    try:
        all_students = []

        for turma in turmas:
            path = os.path.join("Downloads", turma, "students.json")
            students = load_students_from_txt(path)
            submissions_path = os.path.join("Downloads", turma, "submissions")

            rename_files(submissions_path, list_title, questions_data, students)
            save_students_to_txt(students, path)
            all_students.extend(students)

        save_students_to_txt(all_students, os.path.join("Downloads", "students_final.json"))
        log_info("Renomeação e salvamento dos dados finais concluídos com sucesso.")
    except Exception as e:
        log_error(f"Erro ao integrar renomeação no main: {e}")
