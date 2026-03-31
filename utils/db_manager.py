"""
Database Manager
"""

from datetime import datetime
from werkzeug.security import generate_password_hash, check_password_hash
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy.dialects.mysql import LONGBLOB
from enum import Enum

db = SQLAlchemy()

# ENUMS
class JobType(str, Enum):
    TEXT = 'Text'
    FILE = 'File'
    CONVERSATIONAL = 'Conversational'

class JobStatus(str, Enum):
    PROCESSING = 'Processing'
    SUCCESS = 'Success'
    FAILED = 'Failed'

# MODELS
class User(db.Model):
    """User accounts"""
    __tablename__ = 'Users'
    
    Id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    FirstName = db.Column(db.String(100), nullable=False)
    LastName = db.Column(db.String(100), nullable=False)
    Email = db.Column(db.String(255), unique=True, nullable=False, index=True)
    Password = db.Column(db.String(255), nullable=False)
    IsActive = db.Column(db.Boolean, default=True)
    TokenVersion = db.Column(db.Integer, default=0)
    CreatedDate = db.Column(db.DateTime, default=datetime.utcnow)
    ModifiedDate = db.Column(db.DateTime, onupdate=datetime.utcnow)

    jobs = db.relationship('ProcessingJob', backref='user', lazy=True, cascade='all, delete-orphan')
    
    def set_password(self, password: str):
        self.Password = generate_password_hash(password, method='pbkdf2:sha256')
    
    def check_password(self, password: str) -> bool:
        return check_password_hash(self.Password, password)
    
    def invalidate_tokens(self):
        self.TokenVersion += 1
        db.session.commit()
    
    def to_dict(self):
        return {
            'id': self.Id,
            'firstName': self.FirstName,
            'lastName': self.LastName,
            'email': self.Email,
            'isActive': self.IsActive,
            'createdDate': self.CreatedDate.isoformat() if self.CreatedDate else None
        }

class ProcessingJob(db.Model):
    """User's document processing projects (Conversations/Threads)"""
    __tablename__ = 'ProcessingJobs'
    
    Id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    JobName = db.Column(db.String(255), nullable=False)
    IsFavorite = db.Column(db.Boolean, default=False)
    UserId = db.Column(db.Integer, db.ForeignKey('Users.Id', ondelete='CASCADE'), nullable=False, index=True)
    CreatedDate = db.Column(db.DateTime, default=datetime.utcnow)
    
    history = db.relationship('ProcessingJobHistory', backref='job', lazy=True, cascade='all, delete-orphan')
    
    def get_last_activity(self):
        """Used for the sidebar to show the most recent version's summary"""
        return ProcessingJobHistory.query.filter_by(ProcessJobId=self.Id)\
                 .order_by(ProcessingJobHistory.CreatedDate.desc()).first()

    def to_dict(self):
        return {
            'id': self.Id,
            'jobName': self.JobName,
            'userId': self.UserId,
            'createdDate': self.CreatedDate.isoformat() if self.CreatedDate else None,
            'processCount': len(self.history)
        }

class ProcessingJobHistory(db.Model):
    """Individual processing instances (Messages/Versions)"""
    __tablename__ = 'ProcessingJobsHistory'
    
    Id = db.Column(db.Integer, primary_key=True)
    ProcessJobId = db.Column(db.Integer, db.ForeignKey('ProcessingJobs.Id'), nullable=False)
    JobType = db.Column(db.Enum(JobType), nullable=False)
    
    UploadFileName = db.Column(db.String(255))
    UploadFileData = db.Column(LONGBLOB, nullable=True) 
    OutputFileName = db.Column(db.String(255))
    OutputFileData = db.Column(LONGBLOB, nullable=True)
    
    FontFamily = db.Column(db.String(50), default='Calibri')
    FontSize = db.Column(db.Integer, default=11)
    IncludeCover = db.Column(db.Boolean, default=False)
    IncludeTOC = db.Column(db.Boolean, default=False)
    
    ProcessingCount = db.Column(db.Integer, default=1, nullable=False)
    Summary = db.Column(db.Text, nullable=True)
    LastEditedDate = db.Column(db.DateTime, nullable=True)
    
    Status = db.Column(db.Enum(JobStatus), default=JobStatus.PROCESSING)
    ErrorMessage = db.Column(db.Text)
    
    CreatedDate = db.Column(db.DateTime, default=datetime.utcnow)
    ModifiedDate = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    def to_dict(self, include_file_data=False):
        """Serializes history for the chat window UI"""
        data = {
            'id': self.Id,
            'jobId': self.ProcessJobId,
            'type': self.JobType.value if hasattr(self.JobType, 'value') else self.JobType,
            'summary': self.Summary,
            'processingCount': self.ProcessingCount,
            'fontFamily': self.FontFamily,
            'fontSize': self.FontSize,
            'includeCover': self.IncludeCover,
            'includeTOC': self.IncludeTOC,
            'status': self.Status.value if hasattr(self.Status, 'value') else self.Status,
            'errorMessage': self.ErrorMessage,
            'timestamp': self.CreatedDate.strftime("%I:%M %p") if self.CreatedDate else '',
            'createdDate': self.CreatedDate.isoformat() if self.CreatedDate else None,
            'uploadFileName': self.UploadFileName,
            'downloadUrl': f'/api/process/download/{self.Id}'
        }

        if include_file_data and self.UploadFileData:
            job_type_val = self.JobType.value if hasattr(self.JobType, 'value') else self.JobType
            if job_type_val == 'Text':
                try:
                    data['rawText'] = self.UploadFileData.decode('utf-8')
                except (UnicodeDecodeError, AttributeError):
                    data['rawText'] = str(self.UploadFileData)

        return data

# Helper Function
def init_db(app):
    """Initialize database with Flask app"""
    db.init_app(app)
    with app.app_context():
        db.create_all()
        print("✓ Database tables created/updated successfully")