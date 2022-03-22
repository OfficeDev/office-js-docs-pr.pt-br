---
title: Office glossário de termos de complementos
description: Um glossário de termos comumente usado em toda a documentação Office de complementos.
ms.date: 02/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: 327c7a8bcc8c3ab21c437c50003e57d34fb933e0
ms.sourcegitcommit: 4a7b9b9b359d51688752851bf3b41b36f95eea00
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/22/2022
ms.locfileid: "63711214"
---
# <a name="office-add-ins-glossary"></a>Office glossário de complementos

Este é um glossário de termos comumente usados em toda a documentação Office de complementos.

## <a name="add-in"></a>suplemento

Office Os complementos são aplicativos Web que estendem Office aplicativos. Esses aplicativos Web adicionam nova funcionalidade ao aplicativo Office, como trazer dados externos, automatizar processos ou inserir objetos interativos em Office documentos.

Office Os complementos do Office diferem dos complementos VBA, COM e VSTO porque oferecem suporte entre plataformas (geralmente web, Windows, Mac e iPad) e são baseados em tecnologias web padrão (HTML, CSS e JavaScript). A linguagem de programação principal de um Office é JavaScript ou TypeScript.

## <a name="add-in-commands"></a>comandos de add-in

**Comandos de complemento são** elementos de interface do usuário, como botões e menus, que estendem o Office interface do usuário do seu complemento. Quando os usuários selecionam um elemento de comando de complemento, eles iniciam ações como executar código JavaScript ou exibir o complemento em um painel de tarefas. Os comandos de complemento permitem que o seu add-in pareça uma parte do Office, o que dá aos usuários mais confiança no seu complemento. Consulte [Comandos de Excel, PowerPoint Word](../design/add-in-commands.md) e [Add-in](../outlook/add-in-commands-for-outlook.md) para Outlook saber mais.

Consulte também: [faixa de opções, botão faixa de opções](#ribbon-ribbon-button).

## <a name="application"></a>aplicação

**O** aplicativo se refere a um Office aplicativo. Os aplicativos Office que suportam Office de Office são Excel, OneNote, Outlook, PowerPoint, Project e Word.

Consulte também: [cliente](#client), [host](#host), [Office aplicativo, Office cliente](#office-application-office-client).

## <a name="application-specific-api"></a>API específica do aplicativo

APIs específicas do aplicativo fornecem objetos fortemente digitados que interagem com objetos nativos de um aplicativo Office específico. Por exemplo, você chama as APIs Excel JavaScript para acesso a planilhas, intervalos, tabelas, gráficos e muito mais. As APIs específicas do aplicativo estão disponíveis atualmente para Excel, OneNote, PowerPoint, Visio e Word. Consulte [Modelo de API específico do aplicativo](../develop/application-specific-api-model.md) para saber mais.

Consulte também: [API comum](#common-api).

## <a name="client"></a>client

**O** cliente normalmente se refere a um Office aplicativo. Os aplicativos Office, ou clientes, que suportam Office Add-ins são Excel, OneNote, Outlook, PowerPoint, Project e Word.

Consulte também: [aplicativo](#application), [host](#host), [Office aplicativo, Office cliente](#office-application-office-client).

## <a name="common-api"></a>Common API

APIs comuns são usadas para acessar recursos como interface do usuário, caixas de diálogo e configurações de cliente que são comuns em vários Office aplicativos. Esse modelo de API usa [retornos de chamada](https://developer.mozilla.org/docs/Glossary/Callback_function), que permitem especificar apenas uma operação em cada solicitação enviada ao aplicativo do Office.

APIs comuns foram introduzidas com Office 2013 e são usadas para interagir com Office 2013 ou posterior. Algumas APIs comuns são APIs herdas do início de 2010. Excel, PowerPoint e Word têm a funcionalidade da API comum, mas a maior parte dessa funcionalidade foi substituída ou substituída pelo modelo de API específico do aplicativo. As APIs específicas do aplicativo são preferenciais quando possível.

Outras APIs comuns, como as APIs comuns relacionadas a Outlook, interface do usuário e autenticação, são as APIs modernas e preferenciais para essas finalidades. Para obter detalhes sobre o modelo de objeto da API comum, consulte [Common JavaScript API object model](../develop/office-javascript-api-object-model.md).

Consulte também: [API específica do aplicativo](#application-specific-api).

## <a name="content-add-in"></a>add-in de conteúdo

**Os complementos de conteúdo** são webviews, ou exibições do navegador da Web, que são incorporados diretamente Excel, OneNote ou PowerPoint documentos. Os suplementos de conteúdo concedem aos usuários acesso a controles de interface que executam códigos para modificar documentos ou exibir dados de uma fonte de dados. Use suplementos de conteúdo quando quiser inserir a funcionalidade diretamente no documento. Consulte [Content Office Add-ins](../design/content-add-ins.md) para saber mais.

Consulte também: [webview](#webview).

## <a name="content-delivery-network-cdn"></a>rede de distribuição de conteúdo (CDN)

Uma **rede de entrega de** conteúdo **ou CDN** é uma rede distribuída de servidores e data centers. Normalmente, ele fornece maior disponibilidade de recursos e desempenho quando comparado a um único servidor ou data center.

## <a name="contoso"></a>Contoso

**Contoso** Ltd. (também conhecido como Contoso e Contoso University) é uma empresa fictícia usada pela Microsoft como uma empresa e domínio de exemplo.

## <a name="custom-function"></a>função personalizada

Uma **função personalizada** é uma função definida pelo usuário que é empacotada com um Excel de usuário. As funções personalizadas permitem que os desenvolvedores adicionem novas funções, além dos recursos Excel comuns, definindo essas funções em JavaScript como parte de um complemento. Os usuários Excel podem acessar funções personalizadas da mesma forma que qualquer função nativa no Excel. Consulte [Criar funções personalizadas em Excel](../excel/custom-functions-overview.md) para saber mais.

## <a name="custom-functions-runtime"></a>tempo de execução de funções personalizadas

Um **tempo de execução de funções personalizadas** é um tempo de execução JavaScript que executa apenas funções personalizadas. Ela não tem interface do usuário e não pode interagir com Office.js APIs. Se o seu complemento tiver apenas funções personalizadas, esse é um bom tempo de execução leve a ser usado. Se suas funções personalizadas precisam interagir com o painel de tarefas ou Office.js APIs, configure um tempo de execução javaScript compartilhado. Consulte [Configure seu Suplemento do Office para usar em um tempo de execução do JavaScript compartilhado](../develop/configure-your-add-in-to-use-a-shared-runtime.md) para saber mais.

Consulte também: [tempo de execução do JavaScript](#javascript-runtime), [tempo de execução compartilhado do JavaScript, tempo de execução compartilhado](#shared-javascript-runtime-shared-runtime).

## <a name="host"></a>host

**O host** normalmente se refere a um Office aplicativo. Os aplicativos Office ou hosts que suportam Office de Office são Excel, OneNote, Outlook, PowerPoint, Project e Word.

Consulte também: [aplicativo](#application), [cliente](#client), [Office aplicativo, Office cliente](#office-application-office-client).

## <a name="javascript-runtime"></a>Tempo de execução do JavaScript

O **tempo de execução do JavaScript** é o ambiente de host do navegador em que o complemento é executado. No Office no Windows e Office no Mac, o tempo de execução javaScript é um controle de navegador incorporado (ou webview), como Internet Explorer, Edge Legacy, Edge WebView2 ou Safari. Partes diferentes de um complemento executado em tempos de execução javaScript separados. Por exemplo, comandos de complemento, funções personalizadas e código do painel de tarefas normalmente usam tempos de execução JavaScript separados, a menos que você configure um tempo de execução javaScript compartilhado. Consulte [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md) para obter mais informações.

Consulte também: [tempo de execução de funções personalizadas](#custom-functions-runtime), [tempo de execução do JavaScript compartilhado, tempo de execução compartilhado](#shared-javascript-runtime-shared-runtime), [webview](#webview).

## <a name="office-application-office-client"></a>Office aplicativo, Office cliente

**Office cliente refere-se** a um Office aplicativo. Os aplicativos Office, ou clientes, que suportam Office Add-ins são Excel, OneNote, Outlook, PowerPoint, Project e Word.

Consulte também: [aplicativo](#application), [cliente](#client), [host](#host).

## <a name="platform"></a>plataforma

Uma **plataforma** geralmente se refere ao sistema operacional que executa o Office aplicativo. As plataformas que suportam Office de Windows incluem Windows, Mac, iPad e navegadores da Web.

## <a name="quick-start"></a>início rápido

Um **início rápido** é uma descrição de alto nível das principais habilidades e conhecimento necessários para a operação básica de um determinado programa. Na documentação Office de Office, um início rápido é uma introdução ao desenvolvimento de um complemento para um aplicativo específico, como Outlook. Um início rápido contém uma série de etapas que um desenvolvedor de complementos pode concluir em aproximadamente 5 minutos, resultando em um ambiente de desenvolvimento funcional e de complementos funcionando.

Consulte também: [tutorial](#tutorial).

## <a name="requirement-set"></a>conjunto de requisitos

[!include[Requirement set note](../includes/office-js-requirement-sets.md)]

## <a name="ribbon-ribbon-button"></a>faixa de opções, botão de faixa de opções

Uma **faixa** de opções é uma barra de comandos que organiza os recursos de um aplicativo em uma série de guias ou botões na parte superior de uma janela. Um **botão de faixa** de opções é um dos botões desta série. Consulte [Mostrar ou ocultar a faixa de opções Office](https://support.microsoft.com/office/d946b26e-0c8c-402d-a0f7-c6efa296b527#ID0EBBD=Newer_Versions) para obter mais informações.

## <a name="runtime"></a>tempo de execução

Consulte: [Tempo de execução do JavaScript](#javascript-runtime).

## <a name="shared-javascript-runtime-shared-runtime"></a>tempo de execução do JavaScript compartilhado, tempo de execução compartilhado

Um tempo de execução **javaScript** compartilhado ou tempo de execução compartilhado permite que todo o código no seu complemento, incluindo painel de tarefas, comandos de complemento e funções personalizadas, seja executado no mesmo tempo de execução do JavaScript e continue sendo executado mesmo quando o painel de tarefas estiver fechado. Consulte [Configure your Office Add-in to use a shared JavaScript runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md) and [Dicas for using the shared JavaScript runtime in your Office Add-in](https://devblogs.microsoft.com/microsoft365dev/tips-for-using-the-shared-javascript-runtime-in-your-office-add-in%e2%80%af/) to learn more.

Consulte também: [tempo de execução de funções personalizadas](#custom-functions-runtime), [tempo de execução do JavaScript](#javascript-runtime).

## <a name="task-pane"></a>painel de tarefas

Os painéis de tarefas são superfícies de interface ou webviews, que normalmente aparecem no lado direito da janela em Excel, Outlook, PowerPoint e Word. Os painéis de tarefa concedem aos usuários acesso a controles de interface que executam códigos para modificar documentos ou emails ou exibir dados de uma fonte de dados. Use painéis de tarefas quando você não precisa ou não pode inserir a funcionalidade diretamente no documento. Consulte [Painéis de tarefas em Office de complementos](../design/task-pane-add-ins.md) para saber mais.

Consulte também: [webview](#webview).

## <a name="tutorial"></a>tutorial

Um **tutorial** é um auxílio de ensino projetado para ajudar as pessoas a aprender a usar um produto ou procedimento. No contexto Office de Office, um tutorial orienta um desenvolvedor de complementos por meio do processo completo de desenvolvimento de complementos para um aplicativo específico, como Excel. Isso envolve seguir 20 ou mais etapas e é um investimento de tempo maior do que [um início rápido](#quick-start).

Consulte também: [início rápido](#quick-start).

## <a name="ui-less-custom-function"></a>Função personalizada sem interface do usuário

**Funções personalizadas sem interface do usuário são executados** no tempo de execução de funções personalizadas. Eles não têm interface do usuário e não podem interagir com Office.js APIs.

Consulte também: [função personalizada](#custom-function), [tempo de execução de funções personalizadas](#custom-functions-runtime).

## <a name="web-add-in"></a>web add-in

**O complemento da Web** é um termo herdado para um Office Add-in. Esse termo pode ser usado quando Microsoft 365 documentação do Office precisa distinguir os Office modernos de outros tipos de complementos, como VBA, COM ou VSTO.

Consulte também: [add-in](#add-in).

## <a name="webview"></a>webview

Um **webview** é um elemento ou exibição que exibe conteúdo da Web dentro de um aplicativo. Os complementos de conteúdo e os painéis de tarefas contêm navegadores da Web incorporados e são exemplos de webviews em Office de complementos.

Consulte também: [complemento de conteúdo](#content-add-in), [painel de tarefas](#task-pane).

## <a name="xll"></a>XLL

Um **complemento XLL** é um arquivo de Excel que fornece funções definidas pelo usuário e tem a extensão **de arquivo .xll**. Um arquivo XLL é um tipo de arquivo DLL (biblioteca de links dinâmicos) que só pode ser aberto por Excel. Os arquivos de complemento XLL devem ser gravados em C ou C++. Funções personalizadas são o equivalente moderno de funções definidas pelo usuário XLL. As funções personalizadas oferecem suporte em plataformas e são compatíveis com versões anteriores com arquivos XLL. Consulte [Estender funções personalizadas com funções definidas pelo usuário XLL](/office/dev/add-ins/excel/make-custom-functions-compatible-with-xll-udf) para obter mais informações.

Consulte também: [função personalizada](#custom-function).

## <a name="yeoman-generator-yo-office"></a>Gerador Yeoman, yo office

O [gerador Yeoman para Office](../develop/yeoman-generator-overview.md) de usuário usa a ferramenta [Yeoman](https://github.com/yeoman/yo) de código aberto para gerar um Office Add-in por meio da linha de comando. `yo office`é o comando que executa o gerador Yeoman para Office Add-ins. Os Office de complementos rápidos e tutoriais usam o gerador Yeoman.

## <a name="see-also"></a>Confira também

- [Recursos adicionais de suplementos do Office](resources-links-help.md)