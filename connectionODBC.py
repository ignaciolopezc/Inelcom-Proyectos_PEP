import oracledb as oracle


class ConexionODBC:
   
    def __init__(self, host='smt-scan.tchile.local', port=1521, service='explota', user='SRV_MKTB2B', password='Mkt_chile_2025'):
        self.host = host
        self.port = port
        self.service = service
        self.user = user
        self.password = password
        self.connection = None
        self.cursor = None
    
    def conectar(self):
        try:
            # Crear el DSN (Data Source Name)
            dsn = f"{self.host}:{self.port}/{self.service}"
            
            # Establecer la conexión
            self.connection = oracle.connect(
                user=self.user,
                password=self.password,
                dsn=dsn
            )
            
            # Crear cursor
            self.cursor = self.connection.cursor()
            
            print(f"Conexión exitosa a Oracle: {self.host}:{self.port}/{self.service}")
            return True
            
        except oracle.Error as error:
            print(f"Error al conectar a Oracle: {error}")
            return False
    
    def ejecutar_query(self, query):
        if not self.cursor:
            print("Error: No hay conexión activa. Llame a conectar() primero.")
            return None
        
        try:
            self.cursor.execute(query)
            resultados = self.cursor.fetchall()
            print(f"Query ejecutada exitosamente. Filas obtenidas: {len(resultados)}")
            return resultados
        except oracle.Error as error:
            print(f"Error al ejecutar query: {error}")
            return None
    
    def ejecutar_comando(self, comando):
        if not self.cursor:
            print("Error: No hay conexión activa. Llame a conectar() primero.")
            return False
        
        try:
            self.cursor.execute(comando)
            self.connection.commit()
            print(f"Comando ejecutado exitosamente. Filas afectadas: {self.cursor.rowcount}")
            return True
        except oracle.Error as error:
            print(f"Error al ejecutar comando: {error}")
            self.connection.rollback()
            return False
    
    def desconectar(self):
        try:
            if self.cursor:
                self.cursor.close()
                print("Cursor cerrado.")
            
            if self.connection:
                self.connection.close()
                print("Conexión cerrada exitosamente.")
        except oracle.Error as error:
            print(f"Error al cerrar conexión: {error}")
    
    def __enter__(self):
        """
        Permite usar la clase con context manager (with statement).
        """
        self.conectar()
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """
        Cierra la conexión automáticamente al salir del context manager.
        """
        self.desconectar()


# Ejemplo de uso
if __name__ == "__main__":
    # Forma 1: Uso tradicional
    conn = ConexionODBC()
    if conn.conectar():
        resultados = conn.ejecutar_query("SELECT top 4 * FROM TAB_CONS_CHI_FA")
        print(resultados)
        conn.desconectar()
    
    # Forma 2: Usando context manager (recomendado)
    # with ConexionODBC() as conn:
    #     resultados = conn.ejecutar_query("SELECT * FROM dual")
    #     print(resultados)
