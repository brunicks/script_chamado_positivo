import smtplib
import sys
import os
import csv
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime

# Função para enviar o e-mail
def send_email(subject, body, to_email, from_email, smtp_server, smtp_port, smtp_user, smtp_password):
    try:
        print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Configurando o SMTP: servidor={smtp_server}, porta={smtp_port}, usuário={smtp_user}") 
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()  # Inicia o TLS para segurança
        server.login(smtp_user, smtp_password)  # Faz login no servidor SMTP

        # Criação da mensagem
        msg = MIMEMultipart()
        msg['From'] = from_email
        msg['To'] = to_email
        msg['Subject'] = subject

        # Adiciona o corpo da mensagem com HTML
        msg.attach(MIMEText(body, 'html', 'utf-8'))

        # Envia o e-mail
        print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Enviando e-mail para {to_email} com o assunto '{subject}'") 
        server.send_message(msg)
        server.quit()
        return "E-mail enviado com sucesso!"

    except Exception as e:
        return f"Erro ao enviar e-mail: {str(e)}"

def main():
    if len(sys.argv) != 3:
        print("Uso: python send_email.py <file_path> <log_file_path>")
        sys.exit(1)

    file_path = sys.argv[1]
    log_file_path = sys.argv[2]

    # Redireciona saída e erros para um arquivo de log
    with open(log_file_path, 'a', encoding='utf-8') as log_file:
        sys.stdout = log_file
        sys.stderr = log_file

        try:
            # Verifica se o arquivo existe antes de tentar ler
            if not os.path.exists(file_path):
                raise FileNotFoundError(f"O arquivo {file_path} não foi encontrado.")
            
            print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Lendo dados do arquivo: {file_path}") 
            with open(file_path, 'r', encoding='utf-8') as file:
                reader = csv.reader(file, delimiter=',', quotechar='"')
                data = next(reader)
            print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Dados lidos e separados com sucesso.")

            ns, defeito, endereco, complemento, nomecontato, tel, exped = data
            
            # Verifica e lê a assinatura HTML
            signature_file_path = r'C:\caminho\assinatura\html' # Caminho da sua assinatura em html
            if not os.path.exists(signature_file_path):
                raise FileNotFoundError(f"O arquivo de assinatura {signature_file_path} não foi encontrado.")
            with open(signature_file_path, 'r', encoding='utf-8') as signature_file:
                signature_html = signature_file.read()

            # Configura os parâmetros do e-mail
            subject = "Abertura de chamado para a SEFA"
            body = f"""
            <html>
            <body>
                <p>Olá,</p>
                <p>Solicito abertura de chamado para manutenção de máquina na SEFA.</p>
                <p>Desde já agradeço a atenção.</p>
                <p>Número de Série do equipamento Positivo: {ns}<br>
                Defeito apresentado detalhadamente: {defeito}<br>
                Endereço: {endereco}<br>
                Complemento: {complemento}<br>
                Nome da pessoa de contato: {nomecontato}<br>
                Telefone de contato: {tel}<br>
                Horário de expediente do órgão: {exped}</p>
                {signature_html}  <!-- Inclui a assinatura HTML -->
            </body>
            </html>
            """
            to_email = "email@email.com"  # Email do destinatário
            from_email = "email@email.comr"  # Email remetente
            smtp_server = "smtp.office365.com" # Servidor SMTP padrão do outlook
            smtp_port = 587  # Porta padrão para TLS
            smtp_user = "smtp usuario"  # Usuário do SMTP, geralmente é o e-mail
            smtp_password = os.getenv("SMTP_PASSWORD")  # A senha do SMTP deve estar nas variáveis de ambiente

            if smtp_password is None:
                print("Erro: A variável de ambiente SMTP_PASSWORD não está definida.")  
                sys.exit(1)

            # Envia o e-mail
            result = send_email(subject, body, to_email, from_email, smtp_server, smtp_port, smtp_user, smtp_password)
            print(result)

        except Exception as e:
            print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Erro ao ler o arquivo ou enviar o e-mail: {str(e)}")

if __name__ == "__main__":
    main()
