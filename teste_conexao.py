import socket

host = "spsvsql39"
port = 1433

sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
sock.settimeout(5)

print(f"🔌 Testando conexão com {host}:{port}...")

try:
    sock.connect((host, port))
    print(f"✅ Conexão com {host}:{port} bem-sucedida.")
except socket.error as e:
    print(f"❌ Erro ao conectar em {host}:{port} -> {e}")
finally:
    sock.close()
