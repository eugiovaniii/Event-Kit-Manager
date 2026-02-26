# Event Kit Manager

Aplicação desktop em **Python** para gerenciamento de entrega de kits em eventos esportivos.

---

## 📌 Visão Geral

O **Event Kit Manager** é um sistema **offline** desenvolvido para controlar a distribuição de kits de participantes utilizando planilhas do Excel como fonte de dados.

A solução foi criada para ambientes de evento onde:

* Não há acesso à internet
* A velocidade de atendimento é essencial
* A consistência dos dados precisa ser garantida
* Diferentes operadores utilizam cópias separadas da planilha

---

## 🚀 Principais Funcionalidades

* 📥 Importação de planilhas Excel
* 🔎 Busca rápida de participantes por nome
* ✅ Confirmação de entrega com registro automático de data e hora
* 📊 Estatísticas em tempo real de kits entregues
* 💾 Criação automática de backup
* 🔒 Atualização segura de células específicas via mapeamento `EXCEL_ROW`

---

## 🛠 Tecnologias Utilizadas

* **Python 3.10+**
* **Tkinter** — Interface gráfica
* **Pandas** — Manipulação de dados em memória
* **Openpyxl** — Atualização direcionada de células no Excel

---

## ⚙️ Instalação

Instale as dependências:

```bash
pip install pandas openpyxl
```

---

## ▶️ Execução

Execute a aplicação com:

```bash
python app.py
```

---

## 📄 Requisitos da Planilha

A planilha deve conter obrigatoriamente uma aba chamada:

```
GERAL NUMERADA
```

* Os nomes das colunas são normalizados automaticamente pelo sistema.
* O sistema realiza atualizações pontuais nas células, evitando regravação completa do arquivo.

---

## 🏗 Arquitetura

```
Interface Tkinter
        ↓
DataFrame Pandas (em memória)
        ↓
Atualização direcionada com Openpyxl
```

Essa abordagem:

* Evita sobrescrever a planilha inteira
* Reduz riscos de corrupção de dados
* Aumenta a confiabilidade durante o uso ao vivo

---

## 🎯 Caso de Uso

Ideal para operações de entrega de kits em:

* Corridas de rua
* Eventos esportivos
* Competições escolares
* Congressos e credenciamentos

---
