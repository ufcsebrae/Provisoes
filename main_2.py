from sqlalchemy import create_engine

servername = "spsvsql39\\metas"
dbname = "HubDados"
driver = "ODBC+Driver+17+for+SQL+Server"
engine = create_engine(f'mssql+pymssql://@{servername}/{dbname}?trusted_connection=yes&driver={driver}')