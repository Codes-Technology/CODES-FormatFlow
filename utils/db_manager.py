import mysql.connector
import uuid
from datetime import datetime

class DatabaseManager:
    def __init__(self):
        # Update with your root password
        self.config = {
            'user': 'root',
            'password': '$@9176@Codes', 
            'host': '127.0.0.1',
            'database': 'document_api'
        }

    def get_connection(self):
        return mysql.connector.connect(**self.config)

    def create_job(self, file_count: int, original_files_list: list) -> str:
        """Logs a new job with the list of original filenames."""
        job_id = str(uuid.uuid4())
        # Joins the list ['file1.pdf', 'file2.docx'] into one string
        files_string = ", ".join(original_files_list)
        
        conn = self.get_connection()
        cursor = conn.cursor()
        
        # We now include original_files in the INSERT statement
        sql = """
            INSERT INTO processing_jobs (job_id, original_files, status, file_count) 
            VALUES (%s, %s, %s, %s)
        """
        cursor.execute(sql, (job_id, files_string, 'Processing', file_count))
        
        conn.commit()
        cursor.close()
        conn.close()
        return job_id

    def complete_job(self, job_id: str, output_filename: str):
        conn = self.get_connection()
        cursor = conn.cursor()
        sql = "UPDATE processing_jobs SET status = 'Completed', output_filename = %s, completed_at = %s WHERE job_id = %s"
        cursor.execute(sql, (output_filename, datetime.now(), job_id))
        conn.commit()
        cursor.close()
        conn.close()

    def fail_job(self, job_id: str, error_message: str):
        conn = self.get_connection()
        cursor = conn.cursor()
        sql = "UPDATE processing_jobs SET status = 'Failed', error_message = %s WHERE job_id = %s"
        cursor.execute(sql, (error_message, job_id))
        conn.commit()
        cursor.close()
        conn.close()