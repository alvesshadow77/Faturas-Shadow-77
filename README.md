# Faturas Shadow_77

Aplicação local para organizar faturas, comprovativos e lembretes de pagamento. A interface seleciona automaticamente o idioma do Windows (português, inglês, francês, espanhol ou alemão).

## Dados e privacidade

- Os dados da aplicação ficam no perfil Windows do utilizador: `%LOCALAPPDATA%\Faturas Shadow_77`.
- Cada utilizador do Windows configura a sua própria conta; tokens OAuth não são partilhados.
- A app não apaga emails durante o processamento normal. A opção de limpeza Gmail é manual, arquiva primeiro a mensagem e preserva mensagens com estrela.
- A palavra-passe específica de app iCloud é guardada no gestor de credenciais do sistema operativo, não em ficheiros do projeto.
- Nunca publiques ficheiros OAuth, tokens, bases de dados, folhas Excel, faturas, emails, relatórios ou palavras-passe.

## Primeira utilização

Ao abrir pela primeira vez, escolhe Google ou Apple iCloud. A app descarrega anexos PDF não lidos e sincroniza lembretes de faturas por pagar no calendário escolhido.

### Google

O proprietário da aplicação tem de configurar um cliente OAuth do tipo Desktop app no Google Cloud, ativar Gmail API e Google Calendar API e disponibilizar o respetivo ficheiro `credentials.json` localmente. Por omissão, coloca-o em:

```text
%LOCALAPPDATA%\Faturas Shadow_77\config\credentials.json
```

Também pode ser indicado por `FATURAS_GOOGLE_CREDENTIALS`. O utilizador escolhe a conta no ecrã OAuth; o token fica apenas no seu perfil Windows. A publicação pública de uma aplicação que usa Gmail requer a configuração e verificação OAuth aplicáveis no Google Cloud.

### Apple iCloud

Usa o Apple ID e uma palavra-passe específica de app gerada na conta Apple — nunca a palavra-passe normal. A app valida Mail e Calendar antes de guardar a conta. As mensagens não são apagadas.

## Executar a partir do código

```powershell
py -m pip install -r requirements.txt
py scripts\faturas_shadow_77.py
```

Para usar outra pasta de dados, define `FATURAS_SHADOW_77_DIR`. O executável Windows e o instalador são gerados localmente e não devem conter dados de utilizadores.

## Desenvolvimento

```powershell
py -m pytest tests -q
pyinstaller --noconfirm --clean --onefile --console --name "Faturas Shadow_77" --add-data "assets\shadow_77_logo.jpg;assets" scripts\faturas_shadow_77.py
```

O Tesseract OCR é opcional para PDFs só com imagem. Se for usado, tem de ser instalado separadamente.
