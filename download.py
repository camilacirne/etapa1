import os
import io
import re
import time
import shutil
import string
import zipfile
import rarfile
import gspread
import requests
import subprocess
from bs4 import BeautifulSoup
from datetime import datetime
from urllib.parse import urlparse
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from gspread.exceptions import WorksheetNotFound
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.http import MediaIoBaseDownload
from google_auth_oauthlib.flow import InstalledAppFlow
from oauth2client.service_account import ServiceAccountCredentials

SCOPES = [
    "https://www.googleapis.com/auth/classroom.student-submissions.students.readonly",
    "https://www.googleapis.com/auth/classroom.profile.emails",
    "https://www.googleapis.com/auth/classroom.rosters.readonly",
    "https://www.googleapis.com/auth/classroom.courses.readonly",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets"]

def get_credentials():
    try:
        creds = None
        
        if os.path.exists("token.json"):
            creds = Credentials.from_authorized_user_file("token.json", SCOPES)
        
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
                creds = flow.run_local_server(port=0)
        
            with open("token.json", "w") as token:
                token.write(creds.to_json())
        return creds
    except Exception as e:
        log_error(f"Erro em pegar as credenciais: {str(e)}")

def list_classroom_data(service):
    try:
        while True:
            print("\nEscolha a turma:")
            try:
                results = service.courses().list().execute()
                courses = results.get("courses", [])
                
                pif_courses = [course for course in courses if 'PIF' in course['name']]
                
                if not pif_courses:
                    print("Nenhuma turma de PIF encontrada.")
                else:
                    for index, course in enumerate(pif_courses, start=1):
                        print(f"{index} - {course['name']}")
                    print(f"{len(pif_courses) + 1} - Sair")
                    
                    choice = int(input("\nEscolha um número para selecionar a turma: ").strip())
                    if choice == len(pif_courses) + 1:
                        print("Saindo da lista de opções.")
                        return None, None, None, None, None
                    
                    if 1 <= choice <= len(pif_courses):
                        classroom_id = pif_courses[choice - 1]['id']
                        classroom_name = pif_courses[choice - 1]['name']
                    else:
                        print("Opção inválida. Tente novamente.")
                        continue

            except HttpError as error:
                print(f"Ocorreu um erro ao listar as turmas: {error}")
                continue

            print("\nEscolha a lista de exercícios:")
            try:
                assignments = service.courses().courseWork().list(courseId=classroom_id).execute()
                course_work = assignments.get("courseWork", [])
                
                if not course_work:
                    print(f"Nenhuma lista de exercícios encontrada para esta turma: {classroom_name}")
                else:
                    valid_assignments = []
                    for assignment in course_work:
                        title = assignment['title']
                        
                        if any(keyword in title for keyword in ['LISTA', 'LISTAS']):
                            valid_assignments.append(assignment)
                    
                    if not valid_assignments:
                        print(f"Nenhuma lista de exercícios válida encontrada para esta turma: {classroom_name}")
                        return None, None, None, None, None

                    valid_assignments = valid_assignments[::-1]
                    for index, assignment in enumerate(valid_assignments):
                        print(f"{index} - {assignment['title']}")
                    print(f"{len(valid_assignments)} - Sair")
                    
                    choice = int(input("\nEscolha um número para selecionar a lista de exercícios: ").strip())
                    if choice == len(valid_assignments):
                        print("Saindo da lista de opções.")
                        return None, None, None, None, None
                    
                    if 0 <= choice < len(valid_assignments):
                        coursework_id = valid_assignments[choice]['id']
                        list_title = valid_assignments[choice]['title']
                    
                        if 'LISTA' in list_title:
                            list_name = list_title.split(' - ')[0]
                        else:
                            list_name = ' '.join(list_title.split(' - ')[:-1]) 
                    else:
                        print("Opção inválida. Tente novamente.")
                        continue

            except HttpError as error:
                print(f"Um erro ocorreu ao listar as listas de exercícios: {error}")
                continue
            
            return classroom_id, coursework_id, classroom_name, list_name, list_title
    except Exception as e:
        log_error(f"Erro em listar dados do classroom: {str(e)}")    
    
def download_submissions(classroom_service, drive_service, submissions, download_folder, classroom_id, coursework_id, worksheet):
    try:
        print("\nComeçou o download...")
        if worksheet is not None:
            existing_records = worksheet.get_all_values()  
            alunos_registrados = set() 

            for record in existing_records:
                if len(record) > 2: 
                    alunos_registrados.add(record[2])  

        for submission in submissions.get('studentSubmissions', []):
            try:
                student_id = submission['userId']
                student = classroom_service.courses().students().get(courseId=classroom_id, userId=student_id).execute()
                student_email = student['profile']['emailAddress']
                student_login = extract_prefix(student_email)
                student_name = student['profile']['name']['fullName']
                entregou = 1 
                atrasou = 0
                formatacao = 0
                comentarios = None
                copia = 0
                state = submission.get('state', 'UNKNOWN')

                due_date = get_due_date(classroom_service, classroom_id, coursework_id)
                submission_date = get_submission_timestamp(submission, student_id)

                attachments = submission.get('assignmentSubmission', {}).get('attachments', [])
                student_folder = None

                log_info(f"Due date: {due_date}, Submission date: {submission_date}, State: {state}")
                
                if not attachments: 
                    entregou = 0
                    log_info("Aluno não entregou submissão.")
                else:  
                    atrasou = calculate_delay(due_date, submission_date) if due_date and submission_date else 0
                

                if attachments:
                    for attachment in attachments:
                        try:
                            file_id = attachment.get('driveFile', {}).get('id')
                            file_name = attachment.get('driveFile', {}).get('title')
                            request = drive_service.files().get_media(fileId=file_id)

                            file_metadata = drive_service.files().get(fileId=file_id, fields='id, name').execute()
                            if not file_metadata:
                                log_info(f"Não foi possível recuperar os metadados para o arquivo {file_name} de {student_name}.")
                                continue  

                            if file_name.endswith('.c'):
                                if student_folder is None:
                                    student_folder = os.path.join(download_folder, student_login)
                                    if not os.path.exists(student_folder):
                                        os.makedirs(student_folder)
                                        entregou = 1
                                        comentarios="Erro de submissão: enviou arquivo(s), mas não enviou numa pasta compactada."
                                file_path = os.path.join(student_folder, file_name)
                            else:
                                file_path = os.path.join(download_folder, file_name)

                            with io.FileIO(file_path, 'wb') as fh:
                                downloader = MediaIoBaseDownload(fh, request)
                                done = False
                                while not done:
                                    status, done = downloader.next_chunk()
                                    progress_percentage = int(status.progress() * 100)
                                    log_info(f"Baixando {file_name} de {student_name}: {progress_percentage}%")
                                
                                if progress_percentage == 0:
                                    comentarios = "Erro de submissão ou submissão não foi baixada."
                                    entregou = 0
                                    os.remove(file_path)

                        except HttpError as error:
                            if error.resp.status == 403 and 'cannotDownloadAbusiveFile' in str(error):
                                comentarios = "Erro de submissão."
                                if worksheet is not None and student_login not in alunos_registrados:
                                    worksheet.append_rows([[student_name, student_email, student_login,  0, atrasou, formatacao,copia,None, "Erro de submissão: malware ou spam."]])
                                    alunos_registrados.add(student_login)
                                log_info(f"O arquivo {file_name} de {student_name} foi identificado como malware ou spam e não pode ser baixado.")
                                if os.path.exists(file_path):
                                    os.remove(file_path)
                            else:
                                log_info(f"Erro ao baixar arquivo {file_name} de {student_name}: {error}")
                                if worksheet is not None and student_login not in alunos_registrados:
                                    worksheet.append_rows([[student_name, student_email, student_login,  0, atrasou, formatacao,copia,None, "Erro de submissão ou submissão não foi baixada"]])
                                    alunos_registrados.add(student_login)
                            continue  

                        if file_name.endswith('.zip'):
                            expected_name = student_login + '.zip'
                            if file_name != expected_name:
                                corrected_path = os.path.join(download_folder, expected_name)
                                os.rename(file_path, corrected_path)
                                log_info(f"Renomeado {file_name} para {expected_name} de {student_name}.")
                                if file_name.lower() != expected_name: 
                                    formatacao = 0
                                    comentarios = f"Erro de formatação de zip: renomeado {file_name} para {expected_name}. "
                        
                        if file_name.endswith('.rar'):
                            formatacao = 0
                            comentarios = f"Erro de formatação de zip: Foi enviado um rar {file_name}."
                            expected_name = student_login + '.rar'
                            if file_name != expected_name:
                                corrected_path = os.path.join(download_folder, expected_name)
                                os.rename(file_path, corrected_path)
                                log_info(f"Renomeado {file_name} para {expected_name} de {student_name}.")
                                if file_name.lower() != expected_name: 
                                    formatacao = 0
                                    comentarios = f"Erro de formatação de zip: renomeado {file_name} para {expected_name}. "
                                
                else:               
                    log_info(f"Nenhum anexo encontrado para {student_name}")
                    atrasou = 0
                    entregou = 0


                if worksheet is not None and student_login not in alunos_registrados:
                    worksheet.append_rows([[student_name, student_email, student_login, entregou, atrasou, formatacao,copia,None, comentarios]])
                    alunos_registrados.add(student_login)
                
            except Exception as e:
                log_info(f"Erro ao processar {student_name}: {e}")
                if worksheet is not None and student_login not in alunos_registrados:
                    worksheet.append_rows([[student_name, student_email, student_login, 0, atrasou, formatacao,copia,None, "Erro de submissão: erro ao processar."]])
                    alunos_registrados.add(student_login)
                continue  

    except Exception as e:
        log_error(f"Erro ao baixar submissões: {str(e)}")

def extract_prefix(email):
    try:
        return email.split('@')[0]
    except Exception as e:
        log_error(f"Erro em extrair o prefixo do email {str(e)}")
 
    
def get_submission_timestamp(submission, student_id):
    try:
        log_info(f"\nHistórico de submissão: {submission.get('submissionHistory', [])}")
        
        last_timestamp = None  
        
        for history_entry in submission.get('submissionHistory', []):
            state_history = history_entry.get('stateHistory', {})
            state = state_history.get('state')
            actor_user_id = state_history.get('actorUserId')
            timestamp = state_history.get('stateTimestamp')
                      
            if state == 'TURNED_IN' and actor_user_id == student_id:
                last_timestamp = timestamp
        
        return last_timestamp  
    except Exception as e:
        log_error(f"Erro ao calcular a data de submissão pelo estado de entregue: {e}")
        return None
    
def calculate_delay(due_date_str, submission_date_str):
    try:
        due_date = datetime.strptime(due_date_str, "%Y-%m-%dT%H:%M:%S.%fZ")
        submission_date = datetime.strptime(submission_date_str, "%Y-%m-%dT%H:%M:%S.%fZ")
        delay = submission_date - due_date
        if delay.total_seconds() > 0:
         
            delay_in_days = delay.days + 1
            return delay_in_days
        else:
            return 0
    except Exception as e:
        log_error(f"Erro ao calcular atraso: {e}")
        return 0

def get_due_date(classroom_service, classroom_id, coursework_id):
    try:
        coursework = classroom_service.courses().courseWork().get(courseId=classroom_id, id=coursework_id).execute()
        due_date = coursework.get('dueDate')
        due_time = coursework.get('dueTime')

        if due_date:
            year = due_date['year']
            month = due_date['month']
            day = due_date['day']
            
            if due_time:
                hours = due_time.get('hours', 2) # Time Zone GMT 
                minutes = due_time.get('minutes', 59)
                seconds = due_time.get('seconds', 59)
            else:
                hours, minutes, seconds = 2, 59, 59 # Time Zone GMT
            
            return f"{year}-{month:02d}-{day:02d}T{hours:02d}:{minutes:02d}:{seconds:02d}.000Z"
        else:
            return None
    except Exception as e:
        log_error(f"Erro ao obter data de entrega: {e}")
        return None

def update_worksheet(worksheet, student_login, entregou=None, formatacao=None):
    try:
        data = worksheet.get_all_values()
        
        for idx, row in enumerate(data):
            if row[2] == student_login: 
                entregou_atual = row[3] if entregou is None else entregou 
                formatacao_atual = row[5] if formatacao is None else formatacao
                
                worksheet.update([[int(entregou_atual)]], f'D{idx+1}')
                worksheet.update([[int(formatacao_atual)]], f'F{idx+1}')
                
                return
        log_info(f"Login {student_login} não encontrado na planilha.")
    except Exception as e:
        log_error(f"Erro ao atualizar a planilha com formatação ou entrega: {e}")

def update_worksheet_formatacao(worksheet, student_login, formatacao=None, comentario=None):
    try:
        data = worksheet.get_all_values()
        
        for idx, row in enumerate(data):
            if row[2] == student_login: 
                formatacao_atual = row[5] if formatacao is None else formatacao
                comentario_atual = row[6] if comentario is None else comentario

                col = 8
                while col < len(row) and row[col]:   
                    col += 1
                
                cell_range = f'{chr(65+col)}{idx+1}'

                worksheet.batch_update([
                    {
                        'range': f'F{idx+1}',
                        'values': [[int(formatacao_atual)]]
                    },
                    {
                        'range': cell_range,
                        'values': [[comentario_atual]]
                    }
                ])

                return
        log_info(f"Login {student_login} não encontrado na planilha.")
    except Exception as e:
        log_error(f"Erro ao atualizar a planilha com formatação e comentário: {e}")


def update_worksheet_comentario(worksheet, student_login, num_questions=None, comentario=None):
    try:
        data = worksheet.get_all_values()
        
        comentario_index = ord('I') - ord('A') + (num_questions or 0)

        for idx, row in enumerate(data):
            if row[2] == student_login:  

                col = comentario_index

                while col < len(row) and row[col]:   
                    col += 1 
                
                cell_range = f'{chr(65 + col)}{idx + 1}' 
                if comentario:  
                    worksheet.update(values=[[comentario]], range_name=cell_range)
                
                return 
        log_info(f"Login {student_login} não encontrado na planilha.")
    except Exception as e:
        log_error(f"Erro ao atualizar a planilha com comentário: {e}")
