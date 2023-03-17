#Todo move variables to enviroment file to hide from git

PGUSER = 'chappy_pg'
PGPASSWORD = 'GhQrFrDwpwfQKX3'
DATABASE = 'app_tmp'
ip ='localhost'





POSTGRES_URI = f"postgresql://{PGUSER}:{PGPASSWORD}@{ip}/{DATABASE}"