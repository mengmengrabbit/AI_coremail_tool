# Coremail 服务器配置

# 邮箱账户设置
EMAIL_USERNAME = "liym@cffex.com.cn"
EMAIL_PASSWORD = "xxx"

# IMAP 服务器设置
IMAP_SERVER = "mail.cffex.com.cn"  # 使用用户提供的地址
IMAP_PORT = 993  # SSL加密端口
IMAP_USE_SSL = True

# SMTP 服务器设置
SMTP_SERVER = "mail.cffex.com.cn"  # 使用用户提供的地址
SMTP_PORT = 465  # SSL加密端口
SMTP_USE_SSL = True

# 连接设置
CONNECTION_TIMEOUT = 30  # 连接超时时间（秒）
RECONNECT_INTERVAL = 300  # 重连间隔（秒）

# 代理服务器设置
PROXY_ENABLED = True  # 是否启用代理
PROXY_TYPE = "http"  # 代理类型：http 或 socks5
PROXY_HOST = "proxy.cffex.com.cn"  # 代理服务器地址
PROXY_PORT = 8080  # 代理服务器端口
PROXY_USERNAME = ""  # 代理服务器用户名（如果需要认证）
PROXY_PASSWORD = ""  # 代理服务器密码（如果需要认证）

# 邮件获取设置
DEFAULT_MAILBOX = "INBOX"
MAIL_FETCH_LIMIT = 10  # 每次获取的邮件数量限制

# 数据库设置,填写你导出邮件到本地的数据库地址
DATABASE_PATH = "coremail-connect\\data\\email_status.db"

# 邮件文件夹设置，填写你导出邮件到本地的文件夹地址（必填）
EMAIL_FOLDER_PATH = "coremail-connect\\data\\收件箱"

# 日志设置
LOG_LEVEL = "INFO"
LOG_FILE = "coremail_client.log"