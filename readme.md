# SCRIPT DE BLOQUEIO DE CONTAS DO AD


##### Este procedimento tem por finalidade a descrição do script de bloqueio das contas que não estejam efetuando logon na nossa rede por mais de 45 dias.


As contas que obedecerem essas condições serão movidas para a OU de contas desativadas conforme procedimento a seguir:

Acesse o servidor **<hostname>** e vá até o caminho **C:\ScriptBloqueio** e abra o arquivo **Disable_Accounts_AD.vbs**. Este arquivo contém as configurações de onde serão realizadas as ações de bloqueio, movimentação de conta, e envio de log com as contas que foram bloqueadas.

Informe a OU que deverá ser feita a verificação das contas para bloqueio dos usuários conforme exemplo a seguir:
```
strSearchOU="OU=DEP,OU=USUARIOS,OU=BRASIL,OU=SITES"
```
Informe caminho da OU para onde as contas serão movimentadas
```
strNewOU="OU=INATIVOS,OU=USUARIOS,OU=BRASIL,OU=SITES" "
```

Informe o caminho onde irá conter os arquivos de logs gerados no campo **"strLogPath"**.
```
strLogPath="C:\ScriptBloqueio\logs\"
```

Edite o campo **"objEmail.From"** para definir a conta de e-mail de envio.
```
objEmail.From = "BloqueioContas@<dominio>" 
```

Edite o campo **"objEmail.Subject"** para definir o assunto do e-mail.
```
objEmail.Subject = "Rotina - Contas Desabilitadas por inatividade em 45 dias " 
```

Edite o campo **"objEmail.To"** inserindo para qual ou quais e-mails deverão ser enviados os logs com as contas movimentadas conforme a seguir.
```
objEmail.To = "<email>@<dominio>"
```

Edite o campo **"objEmail.Textbody"** para definir o texto do e-mail.
```
objEmail.Textbody = "Rotina - Contas Desabilitadas por inatividade em 45 dias - Vide Anexo!!!"
```

#### NOTA 
Este script não preve autenticação da conta de e-mail, foi habilitado o relay para envio do e-mail.
