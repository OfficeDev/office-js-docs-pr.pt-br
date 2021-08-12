---
title: Visão geral dos suplementos do Excel
description: O suplemento do Excel permite que você estenda a funcionalidade do aplicativo Excel em várias plataformas, como Windows, Mac, iPad e em um navegador.
ms.date: 10/14/2020
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: afca81ad2a4ee3ed24221798f19bea2d1df8dc5f2fe6b97b4d0bcc5152df4783
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57085954"
---
# <a name="excel-add-ins-overview"></a>Visão geral dos suplementos do Excel

Um suplemento do Excel permite que você estenda a funcionalidade do aplicativo Excel em várias plataformas, como Windows, Mac, iPad e em um navegador. Use os suplementos do Excel em uma pasta de trabalho para:

- Interagir com objetos do Excel, ler e gravar dados do Excel.
- Estender a funcionalidade usando o painel de tarefas ou o painel conteúdo baseado na Web
- Adicionar botões personalizados da faixa de opções ou itens de menu contextuais
- Adicionar funções personalizadas
- Fornecer interação mais rica usando janela de caixa de diálogo

A plataforma de Suplementos do Office fornece a estrutura e as APIs JavaScript do Office.js que permitem criar e executar suplementos do Excel. Ao usar a plataforma de Suplementos do Office para criar seu suplemento do Excel, você obterá os seguintes benefícios.

- **Suporte a multiplataformas**: os suplementos do Excel são executados no Office na Web, no Windows, no Mac e no iPad.
- **Implantação centralizada**: os administradores podem implantar rápida e facilmente suplementos do Excel para usuários em toda uma organização.
- **Uso da tecnologia da Web padrão**: Crie um suplemento do Excel usando tecnologias da Web conhecidas, como HTML, CSS e JavaScript.
- **Distribuição pelo AppSource**: Compartilhe o suplemento do Excel com uma ampla audiência publicando-o na [AppSource](https://appsource.microsoft.com/marketplace/apps?product=office&page=1&src=office&corrid=53245fad-fcbe-41f8-9f97-b0840264f97c&omexanonuid=4a0102fb-b31a-4b9f-9bb0-39d4cc6b789d).

> [!NOTE]
> Os suplementos do Excel são diferentes dos suplementos de COM e VSTO, que são anteriores às soluções de integração do Office que são executadas apenas no Office no Windows. Diferentemente dos suplementos de COM, os suplementos do Excel não exigem a instalação de código no dispositivo de um usuário, nem no Excel.

## <a name="components-of-an-excel-add-in"></a>Componentes de um suplemento do Excel

Um suplemento do Excel inclui dois componentes básicos: um aplicativo Web e um arquivo de configuração, chamado de arquivo de manifesto.

O aplicativo Web usa a [API JavaScript do Office](../reference/javascript-api-for-office.md) para interagir com objetos no Excel, e também pode facilitar a interação com recursos online. Por exemplo, um suplemento pode executar qualquer uma das tarefas a seguir.

- Criar, ler, atualizar e excluir dados na pasta de trabalho (planilhas, intervalos, tabelas, gráficos, itens nomeados e muito mais).
- Executar autorização de usuário em um serviço online usando o fluxo padrão OAuth 2.0.
- Emitir solicitações de API ao Microsoft Graph ou qualquer outra API.

O aplicativo Web pode ser hospedado em qualquer servidor Web, além de poder ser criado usando estruturas do lado do cliente (como Angular, React, jQuery) ou tecnologias do lado do servidor (como ASP.NET, Node.js, PHP).

O [manifesto](../develop/add-in-manifests.md) é um arquivo de configuração XML que define como o suplemento integra-se aos clientes do Office, especificando configurações e recursos, como:

- A URL do aplicativo Web do suplemento.
- O nome de exibição, a descrição, a ID, a versão e a localidade padrão do suplemento.
- Como o suplemento integra-se ao Excel, incluindo qualquer interface de usuário personalizada que o suplemento cria (botões da faixa de opções, menus de contexto, etc.).
- Permissões exigidas pelo suplemento, como leitura e gravação no documento.

Para permitir que os usuários finais instalem e usem um suplemento do Excel, você deve publicar o respectivo manifesto no AppSource ou em um catálogo de suplementos. Para obter detalhes sobre como publicar no AppSource, confira [Disponibilizar suas soluções no AppSource e no Office](/office/dev/store/submit-to-appsource-via-partner-center).

## <a name="capabilities-of-an-excel-add-in"></a>Recursos de um suplemento do Excel

Além de interagir com o conteúdo da pasta de trabalho, os suplementos do Excel podem adicionar botões personalizados da faixa de opções ou comandos de menu, inserir painéis de tarefas, adicionar funções personalizadas, abrir caixas de diálogo e, até mesmo, inserir objetos sofisticados baseados na web, como gráficos ou visualizações interativas, em uma planilha.

### <a name="add-in-commands"></a>Comandos de suplemento

Comandos de suplemento são elementos de interface do usuário que estendem a interface do usuário do Excel e iniciam ações no suplemento. É possível adicionar um botão à faixa de opções ou um item a um menu de contexto do Excel. Ao selecionar um comando de suplemento, os usuários iniciam ações como executar código JavaScript ou exibir uma página do suplemento em um painel de tarefas. 

**Comandos de suplemento**

![Comandos de suplemento no Excel.](../images/excel-add-in-commands-script-lab.png)

Para saber mais sobre recursos de comando, plataformas suportadas e práticas recomendadas para o desenvolvimento de comandos de suplemento, confira [Comandos de suplemento para Excel, Word e PowerPoint](../design/add-in-commands.md).

### <a name="task-panes"></a>Painéis de tarefas

Os painéis de tarefas são superfícies de interface que normalmente são exibidas no lado direito da janela no Excel. Os painéis de tarefas dão aos usuários acesso a controles de interface que executam códigos para modificar o documento do Excel ou exibir dados de uma fonte de dados.

**Painel de tarefas**

![Suplemento do painel de tarefas no Excel.](../images/excel-add-in-task-pane-insights.png)

Para saber mais sobre os painéis de tarefas, confira [Painéis de tarefas nos Suplementos do Office](../design/task-pane-add-ins.md). Para ver uma amostra que implementa um painel de tarefas no Excel, confira [Suplemento do Excel JS WoodGrove Expense Trends](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends).

### <a name="custom-functions"></a>Funções personalizadas

Funções personalizadas permitem que desenvolvedores adicionem novas funções do Excel definindo essas funções em JavaScript como parte de um suplemento. Os usuários do Excel podem acessar funções personalizadas da mesma forma que fariam com qualquer função nativa no Excel, como `SUM()`.

**Função personalizada**

![Imagem animada mostrando um usuário final inserindo MYFUNCTION. Função personalizada SPHEREVOLUME em uma célula de uma planilha do Excel.](../images/SphereVolumeNew.gif)

Para obter mais informações sobre funções personalizadas, consulte[Criar funções personalizadas no Excel](custom-functions-overview.md).

### <a name="dialog-boxes"></a>Caixas de diálogo

As caixas de diálogo são superfícies que flutuam acima da janela do aplicativo do Excel ativo. Você pode usar caixas de diálogo para tarefas como exibir páginas de entrada que não podem ser abertas diretamente em um painel de tarefas, solicitar que o usuário confirme uma ação ou hospedar vídeos que possam ser muito pequenos se confinados a um painel de tarefas. Para abrir caixas de diálogo no suplemento do Excel, use a [API da Caixa de Diálogo](/javascript/api/office/office.ui).

**Caixa de diálogo**

![Caixa de diálogo do suplemento no Excel.](../images/excel-add-in-dialog-choose-number.png)

Para saber mais sobre caixas de diálogo e a API da Caixa de Diálogo, confira [Caixas de diálogo nos Suplementos do Office](../design/dialog-boxes.md) e [Usar a API da Caixa de Diálogo em Suplementos do Office](../develop/dialog-api-in-office-add-ins.md).

### <a name="content-add-ins"></a>Suplementos de conteúdo

Os suplementos de conteúdo são superfícies que podem ser inseridas diretamente em documentos do Excel. É possível usar suplementos de conteúdo para inserir objetos sofisticados baseados na Web, como gráficos, visualizações de dados ou mídia em uma planilha ou para conceder aos usuários acesso aos controles de interface que executam código para modificar o documento do Excel ou exibir dados de uma fonte de dados. Use suplementos de conteúdo quando quiser inserir a funcionalidade diretamente no documento.

**Suplemento de conteúdo**

![Suplemento de conteúdo no Excel.](../images/excel-add-in-content-map.png)

Para saber mais sobre suplementos conteúdos, confira [Suplementos do Office de conteúdo](../design/content-add-ins.md). Para ver um exemplo que implementa um suplemento de conteúdo no Excel, confira [Suplemento de conteúdo do Excel Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance) no GitHub.

## <a name="javascript-apis-to-interact-with-workbook-content"></a>APIs JavaScript para interagir com o conteúdo da pasta de trabalho

Um suplemento do Excel interage com objetos no Excel usando a [API JavaScript do Office](../reference/javascript-api-for-office.md), que inclui dois modelos de objetos JavaScript:

- **API JavaScript do Excel**: Introduzida com o Office 2016, a [API JavaScript do Excel](../reference/overview/excel-add-ins-reference-overview.md) fornece objetos do Excel fortemente tipados que você pode usar para acessar planilhas, intervalos, tabelas, gráficos e muito mais.

- **APIs Comuns**: Introduzida com o Office 2013, a API Comum permite que você acesse recursos como interface de usuário, caixas de diálogo e configurações de cliente, que são comuns entre vários tipos de aplicativos do Office. Como a API compartilhada fornece funcionalidade limitada para interação do Excel, você poderá usá-la se seu suplemento precisa ser executado no Excel 2013.

## <a name="next-steps"></a>Próximas etapas

Introdução à [criação de seu primeiro suplemento do Excel](../quickstarts/excel-quickstart-jquery.md). Em seguida, saiba mais sobre os [principais conceitos](excel-add-ins-core-concepts.md) da criação de suplementos do Excel.

## <a name="see-also"></a>Confira também

- [Visão geral da plataforma Suplementos do Office](../overview/office-add-ins.md)
- [Saiba mais sobre Programa para Desenvolvedores do Microsoft 365](https://developer.microsoft.com/microsoft-365/dev-program)
- [Desenvolvimento de Suplementos do Office ](../develop/develop-overview.md)
- [Modelo de objeto JavaScript do Excel em Suplementos do Office](excel-add-ins-core-concepts.md)
- [Referência da API JavaScript do Excel](../reference/overview/excel-add-ins-reference-overview.md)
