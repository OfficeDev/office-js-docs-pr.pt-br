---
title: Visão geral dos suplementos do Excel
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 6d67cd05e7021c8b76fb69135000d4079024713a
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/12/2018
ms.locfileid: "25505955"
---
# <a name="excel-add-ins-overview"></a>Visão geral dos suplementos do Excel

Um suplemento do Excel permite estender a funcionalidade do aplicativo do Excel em várias plataformas, incluindo o Office para Windows, Office Online, Office para o Mac e Office para o iPad. Use suplementos do Excel em uma pasta de trabalho para:

- Interagir com objetos do Excel, ler e gravar dados do Excel. 
- Estender a funcionalidade usando o painel de tarefas ou o painel de conteúdo baseados na web 
- Adicionar botões de faixa de opções personalizados ou itens de menu contextuais
- Proporcionar uma interação mais rica usando a janela de diálogo 

A plataforma de Suplementos do Office fornece a estrutura e as APIs JavaScript Office.js que permitem criar e executar suplementos do Excel. Usando a plataforma de Suplementos do Office para criar seu suplemento do Excel, você terá os seguintes benefícios:

* **Suporte à plataforma cruzada**: Os suplementos do Excel são executados no Office para Windows, Mac, iOS e Office Online.
* **Implantação centralizada**: Os administradores podem implantar rápida e facilmente suplementos do Excel para usuários em toda uma organização.
* **Logon único (SSO)**: Integre facilmente seu suplemento do Excel ao Microsoft Graph.
* **Uso de tecnologia da web padrão**: Crie um suplemento do Excel usando tecnologias da web conhecidas, como HTML, CSS e JavaScript.
* **Distribuição pelo AppSource**: Compartilhe o suplemento do Excel com um público amplo publicando-o no [AppSource](https://appsource.microsoft.com/marketplace/apps?product=office&page=1&src=office&corrid=53245fad-fcbe-41f8-9f97-b0840264f97c&omexanonuid=4a0102fb-b31a-4b9f-9bb0-39d4cc6b789d).

> [!NOTE]
> Suplementos do Excel são diferentes dos suplementos VSTO e COM, que são soluções de integração anteriores do Office executadas apenas no Office para Windows. Ao contrário dos suplementos COM, os suplementos do Excel não exigem a instalação de nenhum código em um dispositivo do usuário ou dentro do Excel. 

## <a name="components-of-an-excel-add-in"></a>Componentes de um suplemento do Excel 

Um suplemento do Excel inclui dois componentes básicos: um aplicativo da web e um arquivo de configuração, chamado de arquivo de manifesto. 

O aplicativo da web usa a [API JavaScript para Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js) para interagir com objetos no Excel, e também pode facilitar a interação com recursos online. Por exemplo, um suplemento pode executar qualquer uma das seguintes tarefas:

* Criar, ler, atualizar e excluir dados na pasta de trabalho (planilhas, intervalos, tabelas, gráficos, itens nomeados e muito mais).
* Executar a autorização do usuário em um serviço online usando o fluxo padrão do OAuth 2.0.
* Emitir solicitações de API ao Microsoft Graph ou qualquer outra API.

O aplicativo da web pode ser hospedado em qualquer servidor web, além de poder ser criado usando estruturas do lado do cliente (como Angular, React, jQuery) ou tecnologias do lado do servidor (como ASP.NET, Node.js, PHP).

O [manifesto](../develop/add-in-manifests.md) é um arquivo de configuração XML que define como o suplemento integra-se aos clientes do Office, especificando configurações e recursos, como: 

* A URL do aplicativo da web do suplemento.
* O nome de exibição, a descrição, a ID, a versão e a localidade padrão do suplemento.
* Como o suplemento integra-se ao Excel, incluindo qualquer interface de usuário personalizada que o suplemento crie (botões da faixa de opções, menus de contexto, etc.).
* Permissões exigidas pelo suplemento, como leitura e gravação no documento.

Para permitir que os usuários finais instalem e usem um suplemento do Excel, você deve publicar o respectivo manifesto no AppSource ou em um catálogo de suplementos. 

## <a name="capabilities-of-an-excel-add-in"></a>Recursos de um suplemento do Excel

Além de interagir com o conteúdo da pasta de trabalho, os suplementos do Excel podem adicionar botões da faixa de opções ou comandos de menu personalizados, inserir painéis de tarefas, abrir caixas de diálogo e até incorporar objetos ricos da web, como gráficos ou visualizações interativas em uma planilha.

### <a name="add-in-commands"></a>Comandos de suplemento

Comandos de suplemento são elementos da interface do usuário que estendem a interface do usuário do Excel e iniciam ações no seu suplemento. Você pode usar comandos de suplemento para adicionar um botão na faixa de opções ou um item a um menu de contexto no Excel. Quando os usuários selecionam um comando de suplemento, iniciam ações como a execução do código JavaScript ou a exibição de uma página do suplemento em um painel de tarefas. 

**Comandos de suplemento**

![Comandos de suplemento no Excel](../images/excel-add-in-commands-script-lab.png)

Para saber mais sobre recursos de comando, plataformas suportadas e práticas recomendadas para o desenvolvimento de comandos de suplemento, confira [Comandos de suplemento para Excel, Word e PowerPoint](../design/add-in-commands.md).

### <a name="task-panes"></a>Painéis de tarefas

Os painéis de tarefas são superfícies de interface que geralmente aparecem no lado direito da janela no Excel. Painéis de tarefas fornecem aos usuários acesso a controles de interface que executam código para modificar o documento do Excel ou exibir dados de uma fonte de dados. 

**Painel de tarefas**

![Suplemento do painel de tarefas no Excel](../images/excel-add-in-task-pane-insights.png)

Para saber mais sobre os painéis de tarefas, confira [Painéis de tarefas nos Suplementos do Office](../design/task-pane-add-ins.md). Para ver uma amostra que implementa um painel de tarefas no Excel, confira [Suplemento JS WoodGrove Expense Trends do Excel](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends).

### <a name="dialog-boxes"></a>Caixas de diálogo

Caixas de diálogo são superfícies que flutuam sobre a janela ativa do aplicativo Excel. Você pode usar caixas de diálogo para tarefas como exibição de páginas de entrada que não podem ser abertas diretamente em um painel de tarefas, solicitar que o usuário confirme uma ação ou vídeos de hospedagem que podem ser muito pequenos se confinados a um painel de tarefas. Para abrir caixas de diálogo em seu suplemento do Excel, use a [API de diálogo](https://docs.microsoft.com/javascript/api/office/office.ui?view=office-js).

**Caixa de diálogo**

![Suplemento de caixa de diálogo no Excel](../images/excel-add-in-dialog-choose-number.png)

Para saber mais sobre caixas de diálogo e a API da Caixa de Diálogo, confira [Caixas de diálogo nos Suplementos do Office](../design/dialog-boxes.md) e [Usar a API da Caixa de Diálogo em Suplementos do Office](../develop/dialog-api-in-office-add-ins.md).

### <a name="content-add-ins"></a>Suplementos de conteúdo

Suplementos de conteúdo são superfícies que você pode incorporar diretamente a documentos do Excel. Você pode usar os suplementos de conteúdo para incorporar objetos ricos baseados na web como gráficos, visualizações de dados ou mídia em uma planilha, ou para dar aos usuários acesso aos controles de interface que executam um código para modificar o documento do Excel ou exibir dados de uma fonte de dados. Use suplementos de conteúdo quando desejar incorporar a funcionalidade diretamente no documento.

**Suplemento de conteúdo**

![Suplemento de conteúdo no Excel](../images/excel-add-in-content-map.png)

Para obter mais informações sobre suplementos de conteúdo, confira [Suplementos de conteúdo do Office](../design/content-add-ins.md). Para ver um exemplo que implementa um suplemento de conteúdo no Excel, confira [Suplemento de conteúdo do Excel Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance) no GitHub.

## <a name="javascript-apis-to-interact-with-workbook-content"></a>APIs JavaScript para interagir com o conteúdo da pasta de trabalho

Um suplemento do Excel interage com objetos no Excel usando a [API JavaScript para Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js), que inclui dois modelos de objeto JavaScript:

* **API JavaScript do Excel**: Introduzida com o Office 2016, a [API JavaScript do Excel](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview?view=office-js) fornece objetos do Excel fortemente tipados que você pode usar para acessar planilhas, intervalos, tabelas, gráficos e muito mais. 

* **API compartilhada**: introduzida com o Office 2013, a API compartilhada permite acessar recursos como interface do usuário, caixas de diálogo e configurações de cliente, que são comuns entre vários tipos de aplicativos host como o Word, Excel e PowerPoint. Como a API compartilhada fornece funcionalidade limitada para interação com o Excel, você pode usá-la se seu suplemento precisa ser executado no Excel 2013.

## <a name="next-steps"></a>Próximas etapas

Comece criando [seu suplemento do Excel](excel-add-ins-get-started-overview.md). Em seguida, aprenda sobre os [conceitos básicos](excel-add-ins-core-concepts.md) de criação de suplementos do Excel.

## <a name="see-also"></a>Veja também

- [Visão geral da plataforma de Suplementos do Office](../overview/office-add-ins.md)
- [Práticas recomendadas para o desenvolvimento de Suplementos do Office](../concepts/add-in-development-best-practices.md)
- [Diretrizes de design para Suplementos do Office](../design/add-in-design.md)
- [Conceitos fundamentais de programação com a API JavaScript do Excel](excel-add-ins-core-concepts.md)
- [Referência da API JavaScript do Excel](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview?view=office-js)
