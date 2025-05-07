from googleapiclient.errors import HttpError
from utils import log_error

def list_classroom_data(service):
    try:
        results = service.courses().list().execute()
        courses = results.get("courses", [])
        pif_courses = [course for course in courses if "PIF" in course["name"]]

        if not pif_courses:
            print("Nenhuma turma de PIF encontrada.")
            return None, None, None, None, None

        print("\nEscolha o semestre:")
        for index, course in enumerate(pif_courses, start=1):
            print(f"{index} - {course['name']}")
        print(f"{len(pif_courses) + 1} - Sair")

        choice = int(input("\nDigite o número da turma: ").strip())
        if choice == len(pif_courses) + 1:
            print("Saindo da seleção.")
            return None, None, None, None, None
        if 1 <= choice <= len(pif_courses):
            classroom = pif_courses[choice - 1]
            classroom_id = classroom["id"]
            classroom_name = classroom["name"]
        else:
            print("Opção inválida.")
            return None, None, None, None, None

        print("\nBuscando listas de exercícios...")
        assignments = service.courses().courseWork().list(courseId=classroom_id).execute()
        course_work = assignments.get("courseWork", [])

        valid_assignments = [
            a for a in course_work if any(k in a["title"].upper() for k in ["LISTA", "LISTAS"])
        ]

        if not valid_assignments:
            print("Nenhuma lista de exercícios encontrada.")
            return None, None, None, None, None

        print("\nEscolha a lista de exercícios:")
        for index, assignment in enumerate(valid_assignments):
            print(f"{index} - {assignment['title']}")
        print(f"{len(valid_assignments)} - Sair")

        choice = int(input("\nDigite o número da lista: ").strip())
        if choice == len(valid_assignments):
            print("Saindo da seleção.")
            return None, None, None, None, None
        if 0 <= choice < len(valid_assignments):
            selected = valid_assignments[choice]
            coursework_id = selected["id"]
            list_title = selected["title"]
            list_name = list_title.split(" - ")[0] if " - " in list_title else list_title
        else:
            print("Opção inválida.")
            return None, None, None, None, None

        return classroom_id, coursework_id, classroom_name, list_name, list_title

    except HttpError as http_err:
        log_error(f"Erro na API do Classroom: {http_err}")
        return None, None, None, None, None
    except Exception as e:
        log_error(f"Erro inesperado ao selecionar dados do Classroom: {e}")
        return None, None, None, None, None
