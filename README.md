# 📊 Automação: Processos Gerais - Returns

Macro desenvolvida em **VBA (Visual Basic for Applications)** para automação de processos no **Microsoft Excel** com integração no **SAP ERP**.
O projeto está organizado de forma modular, facilitando manutenção, versionamento e reutilização do código.

---

## Visão Geral

Este repositório contém a versão totalmente revisada e otimizada da **Planilha Reversa**. Todas as macros de logística reversa foram unificadas em um único ambiente, proporcionando **melhor organização, desempenho aprimorado e uma ampla gama de funcionalidades automatizadas**.

---

## Funcionalidades

A versão atual inclui:

- **Alterações Gerais**: OI (Ordens Inversas), Remessa e TR  
- **Realização de Pré-Cálculo**  
- **Criação de Transporte e Remessa**  
- **Cancelar, Reativar e Zerar Ordens Inversas**  
- **Alteração de NFD**  
- **Cancelamento de ZREC**  
- **Alteração de RFQ**  
- **Lançamento de Ocorrência**  
- **Lançamento de Providência**  
- **Busca de Peso**  
- **Busca de Chave de Acesso e MLOG**  
- **Busca de Endereço**

---

## Controle de Acesso

Para evitar uso simultâneo, a planilha possui um **sistema de bloqueio temporário**.  
O acesso é concedido mediante um **código**, com duração de **1 hora**. Após esse período, a planilha bloqueia automaticamente.  

**Código de acesso geral:** `qb7p7Z001UQTwL`  

**Local na rede:** `LOGISTICA\Atendimento ao Pedido\02. DEVOLUCAO\Reversa`

---

## Níveis de Acesso

| Perfil               | Código de Acesso       | Duração |
|---------------------|----------------------|---------|
| Monitoramento/Geral | qb7p7Z001UQTwL       | 1 hora  |
| CDP                 | wxFd4L99wx6O         | 1 hora  |
| Administrativo      | Sob solicitação      | —       |

---

## Instalação e Uso

1. Abra o arquivo `.xlsm` no Excel com macros habilitadas.  
2. Insira o código de acesso quando solicitado.  
3. Navegue pelas macros disponíveis no **menu de Macros** ou nos botões da planilha.  
4. Siga as instruções para cada processo (OI, Transporte, RFQ, etc.).  

> ⚠️ Certifique-se de que as macros estão habilitadas e que a conexão com o SAP esteja ativa caso utilize funcionalidades automatizadas do SAP.

---

## Estrutura dos Arquivos
 
- `build/` – Planilha macro habilitada (`.xlsm`) pronta para uso  
- `README.md` – Instruções e visão geral do projeto  
- `.gitattributes` – Configurações para detecção de linguagem e tratamento de arquivos

---

## Observações

- Desenvolvido para processos de logística reversa integrados ao SAP.  
- Consolida múltiplos scripts em um **workflow único e otimizado**.  
- Possui **gerenciamento de acesso seguro** para evitar conflitos.
