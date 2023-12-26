## Change License Office 365 Users Script Documentation

### Português

#### Visão Geral
O script PowerShell fornecido automatiza o processo de alteração de licenças para usuários do Office 365 por meio da API Microsoft Graph. Ele oferece a flexibilidade de escolher entre adicionar as licenças Microsoft 365 (M365) E3 ou F3 e remover as licenças correspondentes do Office 365 para uma lista especificada de usuários em um arquivo CSV.

#### Pré-requisitos
1. **Módulo PowerShell do Microsoft Graph:**
   - O script verifica a disponibilidade do módulo PowerShell do Microsoft Graph. Se ausente, solicita que o usuário o instale.
   - O módulo Microsoft Graph pode ser instalado com o comando: `Install-Module Microsoft.Graph -Scope CurrentUser`.
   - Para executar o script é necessesário executar o seguinte comando como administrador: `Set-ExecutionPolicy RemoteSigned`
  
2. **Formato do Arquivo CSV**
   - O arquivo CSV deve conter uma coluna com o cabeçalho 'UPN' (User Principal Name), representando as identidades dos usuários. O script utiliza esta coluna para identificar e modificar as licenças dos usuários.

3. **Autenticação do Usuário**
   - Quando o comando Connect-MgGraph é executado, o script solicita que o usuário faça login no Microsoft Graph. O usuário precisa inserir suas credenciais para estabelecer a conexão.

#### Funções do Script

##### Função `ConnectMgGraphModule`
- Verifica a presença do módulo PowerShell do Microsoft Graph e o instala, se necessário.
- Conecta-se ao Microsoft Graph com as permissões necessárias.

##### Função `Choice`
- Solicita ao usuário o caminho do arquivo CSV contendo a lista de usuários a serem modificados.
- Permite ao usuário escolher entre duas opções de modificação de licença:
  - Opção 1: ADICIONAR M365 E3 e REMOVER Office 365 E3
  - Opção 2: ADICIONAR M365 F3 e REMOVER Office 365 F3
  - Opção 3: ADICIONAR Office 365 E1 e REMOVER Office 365 E3
- Chama a função `ChangeLicense` com base na escolha do usuário.

##### Função `ChangeLicense`
- Modifica a licença para cada usuário no arquivo CSV especificado.
- Desassocia a licença atual do Office 365 e associa a licença Microsoft 365 correspondente com base na opção selecionada.

##### Função `Write-Log`
- Registra eventos com carimbos de data/hora, níveis de log e mensagens personalizadas.
- As mensagens de log são anexadas a um arquivo de log (`log_change_licence.log` por padrão) ou exibidas no console.

##### Função `main`
- Orquestra a execução do script, chamando as funções `ConnectMgGraphModule` e `Choice`.

#### Uso
1. **Executar o Script:**
   - Execute o script no PowerShell ISE

2. **Interação:**
   - Ao se conectar ao Microsoft Graph, o usuário deve fazer login usando suas credenciais.
   - O script solicita ao usuário que informe o caminho do arquivo CSV e escolha uma opção de modificação de licença.

3. **Registros:**
   - Os eventos são registrados em um arquivo de log (`log_change_licence.log` por padrão).

4. **Observação:**
   - Certifique-se de que o módulo PowerShell do Microsoft Graph esteja instalado para uma execução bem-sucedida do script.

#### Exemplo
    .\ChangeLicenseScript.ps1

#### Informações adicionais
    - Author: Pedro Rodrigues
    - Contact: cast.pedroh.utic@sebrae.com.br | phsr2001@gmail.com | ophsr.com.br

---
<div style="page-break-after: always;"></div>

### English

#### Overview
The provided PowerShell script automates the process of changing licenses for Office 365 users through the Microsoft Graph API. It offers the flexibility to choose between adding Microsoft 365 (M365) E3 or F3 licenses and removing the corresponding Office 365 licenses for a specified list of users in a CSV file.

#### Prerequisites
1. **Microsoft Graph PowerShell Module:**
   - The script checks for the availability of the Microsoft Graph PowerShell module. If not present, it prompts the user to install it.
   - The Microsoft Graph module can be installed by running: `Install-Module Microsoft.Graph -Scope CurrentUser`.
   - To run the script, you must run the following command as administrator: `Set-ExecutionPolicy RemoteSigned`

2. **CSV File Format**
   - The CSV file should contain a column with the header 'UPN' (User Principal Name), representing the user identities. The script uses this column to identify and modify user licenses.

3. **User Authentication**
   - When the Connect-MgGraph command is executed, the script prompts the user to log in to Microsoft Graph. The user needs to enter their credentials to establish the connection.

#### Script Functions

##### `ConnectMgGraphModule` Function
- Checks for the presence of the Microsoft Graph PowerShell module and installs it if necessary.
- Connects to Microsoft Graph with the required scopes.

##### `Choice` Function
- Prompts the user to provide the path of the CSV file containing the list of users to be modified.
- Allows the user to choose between two license modification options:
  - Option 1: ADD M365 E3 and REMOVE Office 365 E3
  - Option 2: ADD M365 F3 and REMOVE Office 365 F3
  - Option 3: ADD Office 365 E1 and REMOVE Office 365 E3
- Calls the `ChangeLicense` function based on the user's choice.

##### `ChangeLicense` Function
- Modifies the license for each user in the specified CSV file.
- Unassigns the current Office 365 license and assigns the corresponding Microsoft 365 license based on the selected option.

##### `Write-Log` Function
- Logs events with timestamps, log levels, and custom messages.
- Log messages are appended to a log file (`log_change_licence.log` by default) or displayed in the console.

##### `main` Function
- Orchestrates the execution of the script by calling `ConnectMgGraphModule` and `Choice` functions.

#### Usage
1. **Run the Script:**
   - Run the script on PowerShell ISE 

2. **Interaction:**
   - When connecting to Microsoft Graph, the user must log in using their credentials.
   - The script prompts the user to enter the path of the CSV file and choose a license modification option.
3. **Logging:**
   - Events are logged to a log file (`log_change_licence.log` by default).

4. **Note:**
   - Ensure that the Microsoft Graph PowerShell module is installed for successful script execution.

#### Example
    .\ChangeLicenseScript.ps1

#### Additional Information
    - Author: Pedro Rodrigues
    - Contact: cast.pedroh.utic@sebrae.com.br | phsr2001@gmail.com | ophsr.com.br

