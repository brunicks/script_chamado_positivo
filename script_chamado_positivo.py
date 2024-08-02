import smtplib
import sys
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Função para enviar o e-mail
def send_email(subject, body, to_email, from_email, smtp_server, smtp_port, smtp_user, smtp_password):
    try:
        print(f"Configurando o SMTP: servidor={smtp_server}, porta={smtp_port}, usuário={smtp_user}")  # Mensagem de depuração
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
        print(f"Enviando e-mail para {to_email} com o assunto '{subject}'")  # Mensagem de depuração
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

    try:
        # Lê o arquivo com as informações com codificação UTF-8
        print(f"Lendo dados do arquivo: {file_path}")  # Mensagem de depuração
        with open(file_path, 'r', encoding='utf-8') as file:
            data = file.read().strip()

        # Separa os dados do arquivo
        parts = data.split(',')
        if len(parts) != 7:
            raise ValueError("O arquivo deve conter exatamente 7 partes separadas por vírgula.")
        print("Dados lidos e separados com sucesso.")  # Mensagem de depuração

        ns, defeito, endereco, complemento, nomecontato, tel, exped = parts

        # Lê a assinatura HTML do arquivo
        with open(r'C:\Users\est.07545996984\Documents\assbruno.html', 'r', encoding='utf-8') as signature_file:
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
        to_email = "brunicks02@gmail.com"
        from_email = "bruno.mesquita@celepar.sefa.pr.gov.br"
        smtp_server = "smtp.office365.com"
        smtp_port = 587  # Porta padrão para TLS
        smtp_user = "bruno.mesquita@celepar.sefa.pr.gov.br"
        smtp_password = os.getenv("SMTP_PASSWORD")  # A senha do SMTP deve estar na variável de ambiente

        if smtp_password is None:
            print("Erro: A variável de ambiente SMTP_PASSWORD não está definida.")  # Mensagem de depuração
            sys.exit(1)

        # Redireciona saída e erros para um arquivo de log
        with open(log_file_path, 'a', encoding='utf-8') as log_file:
            # Redireciona a saída padrão e erros para o arquivo de log
            sys.stdout = log_file
            sys.stderr = log_file

            # Envia o e-mail
            result = send_email(subject, body, to_email, from_email, smtp_server, smtp_port, smtp_user, smtp_password)
            print(result)

    except Exception as e:
        print(f"Erro ao ler o arquivo ou enviar o e-mail: {str(e)}")

if __name__ == "__main__":
    main()
