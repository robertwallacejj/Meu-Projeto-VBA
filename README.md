# Projeto VBA Excel — Controle de Recebimento e Indução

Solução em **VBA para Excel** focada em **padronização de lançamentos**, **validação de dados** e **redução de erros operacionais** no processo de recebimento/indução. O projeto utiliza **UserForms**, **validações de negócio**, **persistência em planilhas** e um **calendário customizado** para tornar o preenchimento mais rápido, confiável e fácil de manter.

---

## Visão geral

Este projeto foi desenvolvido para registrar informações operacionais a partir de um formulário no Excel, validando regras antes da gravação e centralizando o fluxo de entrada de dados.

O usuário abre o formulário principal, preenche os campos obrigatórios, seleciona a data por meio de um calendário visual e salva o registro. Durante esse processo, o sistema valida:

- se todos os campos obrigatórios foram preenchidos;
- se a data e os horários informados são válidos;
- se o **agente de cargas** está autorizado;
- se a **AWB** existe na base principal;
- se a AWB já possui registro anterior em `fup_recebimento`.

Com base nessas regras, o VBA define o comportamento de gravação, controla o **ID do lançamento** e trata a **observação** de forma opcional ou obrigatória, dependendo do cenário.

---

## Principais funcionalidades

### Formulário principal de lançamento
O `UserForm1` concentra a operação principal do projeto, permitindo o preenchimento estruturado dos dados de indução e recebimento.

### Validação de campos obrigatórios
Antes de salvar, o sistema verifica se todos os campos mínimos foram preenchidos, evitando registros incompletos.

### Validação de data e hora
O projeto valida:

- data de indução;
- horário inicial;
- horário final.

A data e os horários são consolidados em valores de data/hora completos para gravação na planilha.

### Validação de agente autorizado
O nome selecionado no campo `cmbAGENTEDECARGAS` é validado contra a planilha `AUTORIZADOS`.

### Validação de AWB na base principal
A AWB informada é validada na planilha `fup_aduaneiro`. Caso não exista, o sistema bloqueia o lançamento.

### Regra de recorrência por AWB
Ao gravar em `fup_recebimento`, o sistema trata cenários diferentes:

- **primeiro registro da AWB**: grava com ID inicial;
- **AWB já existente sem ID anterior**: grava/atualiza como primeiro ID;
- **AWB já registrada com ID anterior**: exige observação obrigatória e gera novo ID sequencial.

### Calendário customizado
O projeto possui um formulário específico (`frmCalendario`) para seleção visual de data, com navegação entre meses.

### Limpeza rápida do formulário
Após salvar, ou ao clicar no botão de limpar, os campos do formulário são resetados para um novo lançamento.

---

## Estrutura do projeto

```text
Meu-Projeto-VBA-main/
├── README.md
├── Readme - Projeto Vba Excel.docx
├── Módulo2.bas
├── UserForm1.frm
├── UserForm1.frx
├── frmCalendario.frm
├── frmCalendario.frx
└── clsDia.cls
```

### Componentes principais

#### `Módulo2.bas`
Contém o ponto de entrada principal do projeto:

- `AbrirFUP()` → abre o formulário principal `UserForm1`.

#### `UserForm1.frm`
Formulário principal de lançamento. É responsável por:

- inicializar o combo de agentes autorizados;
- abrir o calendário;
- limpar campos;
- validar dados;
- aplicar regras de negócio;
- gravar os dados em `fup_recebimento`.

#### `frmCalendario.frm`
Formulário de calendário customizado para seleção de data:

- exibe o mês atual;
- permite navegar entre meses;
- monta os dias dinamicamente;
- envia a data escolhida para o campo de destino no formulário principal.

#### `clsDia.cls`
Classe auxiliar usada para associar eventos de clique aos botões dinâmicos do calendário.

---

## Fluxo operacional

### 1. Abrir o formulário
O processo começa pela macro:

```vb
AbrirFUP
```

Ela abre o `UserForm1`, que é a tela principal de operação.

### 2. Preencher os dados
O usuário informa:

- AWB;
- agente de cargas;
- data de indução;
- início da indução;
- fim da indução;
- seleção;
- liberados;
- devolução;
- manifestado;
- APAC;
- fiscalização.

### 3. Selecionar a data
Ao clicar no botão de data, o sistema abre o `frmCalendario`, permitindo escolher a data visualmente.

### 4. Validar regras
Ao salvar, o VBA executa as validações de negócio e consistência.

### 5. Gravar na planilha
Se tudo estiver correto, os dados são gravados em `fup_recebimento`, com os formatos e controles necessários.

### 6. Limpar e preparar novo lançamento
Ao final do processo, o sistema limpa os campos e retorna o foco para o campo da AWB.

---

## Regras de negócio implementadas

### Campos obrigatórios
Os seguintes controles são tratados como obrigatórios:

- `txtAWB`
- `cmbAGENTEDECARGAS`
- `txtDataInducao`
- `txtInicioInducao`
- `txtFimInducao`
- `txtSelecao`
- `txtLiberados`
- `txtDevolucao`
- `txtManifestado`
- `txtAPAC`
- `txtFiscalizacao`

### Agentes autorizados
O agente informado deve existir na coluna `A` da planilha `AUTORIZADOS`.

### Validação de AWB
A AWB deve existir na coluna `C` da planilha `fup_aduaneiro`.

### Observação
A observação segue a regra:

- **opcional** no primeiro registro;
- **obrigatória** quando a AWB já possui registro anterior com ID.

### ID do registro
O sistema grava o ID em `fup_recebimento`, coluna `P`, de acordo com o histórico da AWB.

---

## Planilhas esperadas no arquivo Excel

Para que o projeto funcione corretamente, o workbook deve conter pelo menos estas planilhas:

### `AUTORIZADOS`
Usada para carregar e validar os agentes de cargas.

- coluna `A`: lista de agentes autorizados.

### `fup_aduaneiro`
Usada para validar a existência da AWB.

- coluna `C`: AWBs válidas.

### `fup_recebimento`
Usada para gravar os lançamentos processados pelo formulário.

Colunas utilizadas pelo código:

- `A` → AWB
- `B` → Data/hora início
- `E` → Data/hora fim
- `H` → Seleção
- `I` → Liberados
- `J` → Devolução
- `K` → Manifestado
- `L` → APAC
- `M` → Fiscalização
- `N` → Timestamp de gravação (`Now`)
- `O` → Observação
- `P` → ID do registro
- `Q` → Agente de cargas

---

## Como executar

### Pré-requisitos

- Microsoft Excel com suporte a **VBA**;
- macros habilitadas;
- referências padrão do **MSForms** disponíveis no ambiente do Office.

### Passo a passo

1. Abra o arquivo Excel habilitado para macro.
2. Garanta que as planilhas `AUTORIZADOS`, `fup_aduaneiro` e `fup_recebimento` existam.
3. Importe os arquivos `.bas`, `.frm`, `.frx` e `.cls` no editor VBA, caso esteja montando o projeto manualmente.
4. Execute a macro:

```vb
AbrirFUP
```

5. Preencha os campos do formulário e salve o registro.

---

## Experiência do usuário

O projeto já inclui alguns cuidados de usabilidade:

- formulário centralizado na tela;
- calendário visual para seleção de data;
- mensagens de erro orientadas ao usuário;
- foco automático no campo que precisa correção;
- limpeza automática dos campos após salvar;
- lista de agentes carregada automaticamente no `Initialize` do formulário.

---

## Pontos fortes do projeto

- fluxo objetivo e fácil de operar;
- regras de negócio aplicadas antes da gravação;
- validação de agente autorizado e AWB;
- controle de recorrência por AWB;
- calendário customizado sem dependência externa;
- estrutura modular com formulário principal, calendário e classe auxiliar.

---

## Pontos de atenção

Embora o projeto esteja funcional, alguns pontos merecem cuidado em futuras evoluções:

- o nome das planilhas e colunas precisa permanecer consistente com o código;
- a lógica está fortemente vinculada à estrutura do workbook;
- validações numéricas adicionais podem ser úteis, dependendo do processo;
- tratamento de duplicidade pode ser refinado conforme a regra operacional evoluir;
- o projeto ainda pode ganhar melhorias visuais no `UserForm1`.

---

## Sugestões de evolução

Possíveis melhorias futuras:

- padronizar melhor os captions e nomes visuais do formulário;
- adicionar máscara/validação mais forte para hora;
- incluir log de alterações por usuário;
- criar tela de consulta de registros salvos;
- adicionar edição controlada de registros existentes;
- implementar tratamento mais detalhado para exceções operacionais;
- melhorar identidade visual do formulário principal.

---

## Tecnologias utilizadas

- **VBA (Visual Basic for Applications)**
- **Excel UserForms**
- **MSForms Controls**
- **Planilhas Excel como base de dados operacional**

---

## Resumo técnico

O projeto é uma aplicação VBA orientada a formulário, com foco em entrada controlada de dados e validações de negócio. A arquitetura atual está dividida em:

- **macro de abertura**;
- **formulário principal**;
- **calendário dinâmico**;
- **classe de eventos para botões dinâmicos**.

É uma base sólida para automações internas em Excel que exigem preenchimento guiado, consistência operacional e menor dependência de digitação manual.

---



