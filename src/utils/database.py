import sqlite3
import os
from datetime import datetime

class DatabaseManager:
    """数据库管理器 - 处理邮件完成状态的持久化存储"""
    
    def __init__(self, db_path="data/email_status.db"):
        self.db_path = db_path
        # 确保数据库目录存在
        os.makedirs(os.path.dirname(db_path), exist_ok=True)
        self.init_database()
    
    def init_database(self):
        """初始化数据库表"""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            
            # 创建邮件完成状态表
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS email_completion_status (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    application_no TEXT NOT NULL,
                    file_path TEXT NOT NULL,
                    subject TEXT NOT NULL,
                    deadline DATE NOT NULL,
                    completed BOOLEAN DEFAULT FALSE,
                    completed_at TIMESTAMP,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    UNIQUE(application_no, file_path)
                )
            ''')
            
            # 创建索引
            cursor.execute('''
                CREATE INDEX IF NOT EXISTS idx_application_no 
                ON email_completion_status(application_no)
            ''')
            
            cursor.execute('''
                CREATE INDEX IF NOT EXISTS idx_completed 
                ON email_completion_status(completed)
            ''')
            
            conn.commit()
    
    def mark_email_completed(self, application_no, file_path, subject, deadline):
        """标记邮件为已完成"""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            
            cursor.execute('''
                INSERT OR REPLACE INTO email_completion_status 
                (application_no, file_path, subject, deadline, completed, completed_at, updated_at)
                VALUES (?, ?, ?, ?, TRUE, ?, ?)
            ''', (application_no, file_path, subject, deadline, 
                  datetime.now().isoformat(), datetime.now().isoformat()))
            
            conn.commit()
            return cursor.rowcount > 0
    
    def mark_email_uncompleted(self, application_no, file_path):
        """标记邮件为未完成"""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            
            cursor.execute('''
                UPDATE email_completion_status 
                SET completed = FALSE, completed_at = NULL, updated_at = ?
                WHERE application_no = ? AND file_path = ?
            ''', (datetime.now().isoformat(), application_no, file_path))
            
            conn.commit()
            return cursor.rowcount > 0
    
    def is_email_completed(self, application_no, file_path):
        """检查邮件是否已完成"""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            
            cursor.execute('''
                SELECT completed FROM email_completion_status 
                WHERE application_no = ? AND file_path = ?
            ''', (application_no, file_path))
            
            result = cursor.fetchone()
            return result[0] if result else False
    
    def get_completed_emails(self):
        """获取所有已完成的邮件"""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            
            cursor.execute('''
                SELECT application_no, file_path, subject, deadline, completed_at
                FROM email_completion_status 
                WHERE completed = TRUE
                ORDER BY completed_at DESC
            ''')
            
            return cursor.fetchall()
    
    def get_completion_stats(self):
        """获取完成状态统计"""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            
            cursor.execute('''
                SELECT 
                    COUNT(*) as total,
                    SUM(CASE WHEN completed = TRUE THEN 1 ELSE 0 END) as completed,
                    SUM(CASE WHEN completed = FALSE THEN 1 ELSE 0 END) as pending
                FROM email_completion_status
            ''')
            
            result = cursor.fetchone()
            return {
                'total': result[0] if result else 0,
                'completed': result[1] if result else 0,
                'pending': result[2] if result else 0
            }
    
    def cleanup_old_records(self, days=90):
        """清理旧记录（默认90天前的记录）"""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            
            cutoff_date = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
            cutoff_date = cutoff_date.replace(day=cutoff_date.day - days)
            
            cursor.execute('''
                DELETE FROM email_completion_status 
                WHERE created_at < ? AND completed = TRUE
            ''', (cutoff_date.isoformat(),))
            
            conn.commit()
            return cursor.rowcount