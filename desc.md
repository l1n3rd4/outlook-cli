# Conectando ao Amazon RDS PostgreSQL com autenticação IAM

Este guia descreve como estabelecer uma conexão segura a uma instância do Amazon RDS
PostgreSQL utilizando **autenticação via IAM** (token temporário em vez de senha estática).

São apresentados dois caminhos: conexão **via console AWS (interface web)** e conexão
**via CLI (`psql`)** com SSL `verify-full`. Ambos usam os mesmos dados de conexão,
descritos na seção a seguir.

---

## Pré-requisitos

Antes de começar, garanta que você tenha:

- **Login no AWS SSO (IAM Identity Center)** realizado para obter as credenciais de acesso:
  1. Acesse a URL do portal de acesso: https://d-9067824a2a.awsapps.com/start/#/?tab=accounts
  2. Faça login com o seu e-mail Microsoft `@direcional.com.br`.
  3. No menu de contas, clique na conta correspondente ao ambiente que você precisa acessar
     (`dev` | `stg` | `prd`) e selecione a role desejada para carregar as credenciais.

- **AWS CLI** instalada e configurada com credenciais válidas (via `aws configure`,
  variáveis de ambiente ou uma role IAM). O comando `aws rds generate-db-auth-token`
  depende dessas credenciais.
  - [Instruções de instalação da AWS CLI](https://docs.aws.amazon.com/pt_br/cli/latest/userguide/getting-started-install.html#getting-started-install-instructions)

- **Cliente `psql`** (PostgreSQL client) instalado localmente, caso vá usar o caminho via CLI.
- **Permissão IAM** para gerar tokens de autenticação na instância RDS de destino
  (política `rds-db:connect` associada ao usuário/role para o `dbuser` correspondente).
- **Usuário de banco de dados** habilitado para autenticação IAM (fornecido pela equipe de IAM).
- **Conectividade de rede** com a instância RDS (Security Group liberando a porta `5432`,
  rota via VPC/VPN/peering conforme o ambiente).

---

## Dados de conexão (comum às duas opções)

Independentemente de qual caminho você escolher (interface web ou CLI), os três dados
abaixo são necessários. Levante-os **antes** de prosseguir:

| Dado         | Descrição                                                                                      | Exemplo                                          |
|--------------|------------------------------------------------------------------------------------------------|--------------------------------------------------|
| **RDSHOST**  | Endpoint (hostname) da instância RDS.                                                          | `meu-banco.abc123.us-east-1.rds.amazonaws.com`   |
| **DBNAME**   | Nome do banco de dados ao qual você deseja se conectar.                                        | `minha_base`                                     |
| **RDSUSER**  | Usuário de banco de dados habilitado para autenticação IAM, fornecido pela equipe de IAM.      | `iam_user_dev`                                   |

Na **Opção A** (interface), você informará esses valores diretamente nos campos do formulário
do console AWS. Na **Opção B** (CLI), eles serão exportados como variáveis de ambiente.

---

## Opção A — Conexão via interface web (Console AWS)

Se você preferir usar a interface gráfica da AWS em vez da linha de comando, siga os
passos abaixo para se conectar diretamente pelo painel do Amazon RDS.

### Passo 1 — Acessar o painel do Amazon RDS

Após fazer login no portal SSO (URL dos pré-requisitos acima), navegue até o serviço
**Aurora and RDS** no console da AWS.

### Passo 2 — Selecionar o banco de dados

No painel do RDS, clique em **Bancos de dados** (ou **Databases**, se o console estiver
em inglês). Localize e clique no banco de dados ao qual deseja se conectar.

### Passo 3 — Obter os dados de conexão

Na página inicial do banco, você verá as informações de endpoint e porta, semelhante
à tela abaixo:

![screenshot_2026-09-03_154414.png](/screenshot_2026-09-03_154414.png)

### Passo 4 — Gerar a senha (token IAM)

O formulário do console exige uma senha. No caso da autenticação IAM, essa senha é um
**token temporário** gerado via AWS CLI (válido por 15 minutos). Execute o script abaixo
no seu terminal, substituindo os valores entre `< >` pelos
[dados de conexão](#dados-de-conexão-comum-às-duas-opções) do seu ambiente. O token será
impresso na saída — copie-o e cole no campo de senha do formulário no passo seguinte.

```bash
export RDSHOST="<RDS-HOST>.us-east-1.rds.amazonaws.com"  
export RDSUSER=<USER PROVIDENCIADO PELA EQUIPE DE IAM>  
export DBNAME="<NOME DO BANCO>"  
  
aws rds generate-db-auth-token --hostname $RDSHOST --port 5432 --username $RDSUSER --region us-east-1
```

### Passo 5 — Preencher os campos de conexão

Utilize o formulário de conexão do console, mapeando os campos da tabela de
[Dados de conexão](#dados-de-conexão-comum-às-duas-opções) conforme abaixo:

![screenshot_2026-09-03_154031.png](/screenshot_2026-09-03_154031.png)

| Campo no formulário | Valor (referência à tabela acima)                |
|---------------------|--------------------------------------------------|
| **Host**            | `RDSHOST` — endpoint da instância RDS            |
| **Banco de Dados**  | `DBNAME` — nome do banco                         |
| **Nome de usuário** | `RDSUSER` — usuário providenciado pela equipe de IAM |
| **Senha**           | Token IAM gerado no Passo 4                      |

Após preencher, confirme a conexão pela interface.

---

## Opção B — Conexão via CLI (`psql`)

Para quem prefere a linha de comando, o fluxo é composto por três etapas:

1. Baixar o pacote de certificados raiz (CA bundle) da Amazon RDS.
2. Exportar os [dados de conexão](#dados-de-conexão-comum-às-duas-opções) como variáveis
   de ambiente e gerar um token de autenticação temporário via IAM.
3. Abrir a sessão `psql` usando esse token como senha, com verificação total do certificado.

---

### Passo 1 — Baixar o certificado raiz (CA bundle) da Amazon RDS

O `sslmode=verify-full` exige que o cliente valide o certificado apresentado pelo servidor
contra uma autoridade certificadora (CA) confiável. A Amazon publica um pacote global de
certificados raiz que cobre todas as regiões. O comando abaixo baixa esse pacote e o salva
localmente como `global-bundle.pem` (a flag `-o` do `curl` define o nome do arquivo de saída):

```bash
curl -o global-bundle.pem https://truststore.pki.rds.amazonaws.com/global/global-bundle.pem
```

Após executar, você terá um arquivo `global-bundle.pem` no diretório atual, que será
referenciado mais adiante pelo parâmetro `sslrootcert`.

---

### Passo 2 — Exportar variáveis de ambiente e gerar o token IAM

Aqui você exporta os mesmos [dados de conexão](#dados-de-conexão-comum-às-duas-opções) como
variáveis de ambiente e, em seguida, gera um **token de autenticação temporário** que
substitui a senha tradicional. Esse token é válido por um curto período (por padrão,
15 minutos), então a conexão deve ser aberta logo após gerá-lo.

Substitua os valores entre `< >` pelos dados reais do seu ambiente:

```bash
export RDSHOST="<RDS-HOST>.us-east-1.rds.amazonaws.com" 
export RDSUSER=<USER PROVIDENCIADO PELA EQUIPE DE IAM>
export DBNAME="<NOME DO BANCO>"

export RDSPASS=$(aws rds generate-db-auth-token --hostname $RDSHOST --port 5432 --username $RDSUSER --region us-east-1)

psql "host=$RDSHOST port=5432 dbname=$DBNAME user=$RDSUSER sslmode=verify-full sslrootcert=./global-bundle.pem password=$RDSPASS"
```

Detalhamento dos parâmetros de `aws rds generate-db-auth-token`:

- `--hostname $RDSHOST` — endpoint da instância RDS.
- `--port 5432` — porta padrão do PostgreSQL.
- `--username $RDSUSER` — usuário de banco de dados que receberá o token.
- `--region us-east-1` — região AWS onde a instância está hospedada.

---

### Passo 3 — Entendendo a string de conexão do `psql`

O comando `psql` acima abre a sessão interativa usando os parâmetros:

- `host=$RDSHOST` — endpoint da instância.
- `port=5432` — porta do PostgreSQL.
- `dbname=$DBNAME` — banco de dados de destino.
- `user=$RDSUSER` — usuário autenticado via IAM.
- `sslmode=verify-full` — modo de SSL mais rigoroso: criptografa a conexão **e** verifica
  se o certificado do servidor é válido e se o hostname bate com o certificado.
- `sslrootcert=./global-bundle.pem` — caminho para o CA bundle baixado no Passo 1, usado
  para validar o certificado do servidor.
- `password=$RDSPASS` — o token IAM gerado no Passo 2, usado no lugar de uma senha fixa.

---

## Observações finais

- O token IAM é temporário. Se a sessão expirar ou você receber erro de autenticação após
  algum tempo, gere um novo token repetindo o Passo 2 da Opção B.
- Mantenha o arquivo `global-bundle.pem` atualizado periodicamente, pois os certificados
  raiz da Amazon são renovados ao longo do tempo.
- Nunca versione tokens ou credenciais em repositórios; as variáveis de ambiente acima
  vivem apenas na sessão atual do terminal.
