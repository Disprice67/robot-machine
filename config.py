from os import getenv, getcwd

from dotenv import load_dotenv

load_dotenv()

#Main_info
ROOT_DIR = getcwd() + '\SupportFile'
DATABASE = getenv('DATABASE')
DATABASE_TEST = getenv('DATABASE_TEST')
#Ebay
API_KEY = getenv('API_KEY')
CERT_ID = getenv('CERT_ID')
DEV_ID = getenv('DEV_ID')
TOKEN = getenv('TOKEN')
#Sql
ARCHIVE = getenv('ARCHIVE')
SQL_ARCHIVE = getenv('SQL_ARCHIVE')
SQL_ARCHIVE_IN = getenv('SQL_ARCHIVE_IN')
SQL_ARCHIVE_OUT = getenv('SQL_ARCHIVE_OUT')
SQL_COLLISION = getenv('SQL_COLLISION')
SQL_CATEGORY = getenv('SQL_CATEGORY')
SQL_CORPUSE = getenv('SQL_CORPUSE')
SQL_CORPUSE_IN = getenv('SQL_CORPUSE_IN')
SQL_CORPUSE_OUT = getenv('SQL_CORPUSE_OUT') 
SQL_PURCHASE = getenv('SQL_PURCHASE')
SQL_PURCHASE_IN = getenv('SQL_PURCHASE_IN')
SQL_PURCHASE_OUT = getenv('SQL_PURCHASE_OUT')
SQL_PURCHASE_TWO = getenv('SQL_PURCHASE_TWO')
SQL_PURCHASE_TWO_IN = getenv('SQL_PURCHASE_TWO_IN')
SQL_PURCHASE_TWO_OUT = getenv('SQL_PURCHASE_TWO_OUT')
SQL_SHASSIS_IN = getenv('SQL_SHASSIS_IN')
#Credentials
USERNAME_GMAIL = getenv('USERNAME_GMAIL')
PASSWORD_GMAIL = getenv('PASSWORD_GMAIL')
#Mail_server
MAIL_SERVER = getenv('MAIL_SERVER')
#Parser
URL_HUAWEI = getenv('URL_HUAWEI')
USER_AGENT = getenv('USER_AGENT')
