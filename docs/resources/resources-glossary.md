---
title: Office glossário de termos de suplementos
description: Um glossário de termos comumente usados em toda a documentação Office suplementos.
ms.date: 06/15/2022
ms.localizationpriority: medium
ms.openlocfilehash: 002c61cf482da75a5fa2bef0219990ffc9b04034
ms.sourcegitcommit: d8fbe472b35c758753e5d2e4b905a5973e4f7b52
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/25/2022
ms.locfileid: "66229642"
---
# <a name="office-add-ins-glossary"></a>Office glossário de suplementos

Esse é um glossário de termos comumente usados em toda a documentação Office suplementos.

## <a name="add-in"></a>suplemento

Office suplementos são aplicativos Web que estendem Office aplicativos. Esses aplicativos Web adicionam novas funcionalidades ao aplicativo Office, como trazer dados externos, automatizar processos ou inserir objetos interativos em Office documentos.

Office suplementos diferem dos suplementos VBA, COM e VSTO porque oferecem suporte multiplataforma (geralmente Web, Windows, Mac e iPad) e são baseados em tecnologias Web padrão (HTML, CSS e JavaScript). A principal linguagem de programação de um Office suplemento é JavaScript ou TypeScript.

## <a name="add-in-commands"></a>comandos de suplemento

**Os comandos de suplemento** são elementos da interface do usuário, como botões e menus, que estendem a interface do usuário Office para o suplemento. Quando os usuários selecionam um elemento de comando de suplemento, eles iniciam ações como executar código JavaScript ou exibir o suplemento em um painel de tarefas. Os comandos de suplemento permitem que o suplemento se sinta como parte do Office, o que dá aos usuários mais confiança em seu suplemento. Consulte [os comandos de suplemento Excel, PowerPoint word](../design/add-in-commands.md) e suplemento para [Outlook para saber](../outlook/add-in-commands-for-outlook.md) mais.

Consulte também: faixa [de opções, botão da faixa de opções](#ribbon-ribbon-button).

## <a name="application"></a>aplicação

**O** aplicativo refere-se a um Office aplicativo. Os Office que dão suporte Office suplementos são Excel, OneNote, Outlook, PowerPoint, Project e Word.

Consulte também: [cliente](#client), [host](#host), [Office aplicativo, Office cliente](#office-application-office-client).

## <a name="application-specific-api"></a>API específica do aplicativo

As APIs específicas do aplicativo fornecem objetos fortemente tipados que interagem com objetos que são nativos de um aplicativo Office específico. Por exemplo, você chama as Excel APIs JavaScript para acesso a planilhas, intervalos, tabelas, gráficos e muito mais. As APIs específicas do aplicativo estão disponíveis atualmente para Excel, OneNote, PowerPoint, Visio e Word. Consulte [o modelo de API específico do aplicativo](../develop/application-specific-api-model.md) para saber mais.

Consulte também: [API comum](#common-api).

## <a name="client"></a>Cliente

**O** cliente normalmente se refere a um Office aplicativo. Os Office ou clientes que dão suporte Office suplementos são Excel, OneNote, Outlook, PowerPoint, Project e Word.

Consulte também: [aplicativo](#application), [host](#host), [Office aplicativo, Office cliente](#office-application-office-client).

## <a name="common-api"></a>Common API

APIs comuns são usadas para acessar recursos como interface do usuário, caixas de diálogo e configurações de cliente que são comuns em vários Office aplicativos. Esse modelo de API usa [retornos de chamada](https://developer.mozilla.org/docs/Glossary/Callback_function), que permitem especificar apenas uma operação em cada solicitação enviada ao aplicativo do Office.

ApIs comuns foram introduzidas com o Office 2013 e são usadas para interagir com o Office 2013 ou posterior. Algumas APIs comuns são APIs herdadas do início de 2010. Excel, PowerPoint e Word têm funcionalidades comuns de API, mas a maior parte dessa funcionalidade foi substituída ou substituída pelo modelo de API específico do aplicativo. As APIs específicas do aplicativo são preferenciais quando possível.

Outras APIs comuns, como as APIs comuns relacionadas Outlook, interface do usuário e autenticação, são as APIs modernas e preferenciais para essas finalidades. Para obter detalhes sobre o modelo de objeto da API Comum, consulte [o modelo de objeto da API JavaScript comum](../develop/office-javascript-api-object-model.md).

Consulte também: [API específica do aplicativo](#application-specific-api).

## <a name="content-add-in"></a>suplemento de conteúdo

**Os suplementos de conteúdo** são modos de exibição da Web ou exibições do navegador da Web que são inseridos diretamente em Excel, OneNote ou PowerPoint documentos. Os suplementos de conteúdo concedem aos usuários acesso a controles de interface que executam códigos para modificar documentos ou exibir dados de uma fonte de dados. Use suplementos de conteúdo quando quiser inserir a funcionalidade diretamente no documento. Consulte [Suplementos Office conteúdo para](../design/content-add-ins.md) saber mais.

Consulte também: [modo de exibição da Web](#webview).

## <a name="content-delivery-network-cdn"></a>rede de distribuição de conteúdo (CDN)

Uma **rede de distribuição de** **conteúdo ou CDN** é uma rede distribuída de servidores e data centers. Normalmente, ele fornece maior disponibilidade e desempenho de recursos quando comparado a um único servidor ou data center.

## <a name="contoso"></a>Contoso

**Contoso** Ltd. (também conhecida como Contoso e Contoso University) é uma empresa fictícia usada pela Microsoft como uma empresa e domínio de exemplo.

## <a name="custom-function"></a>função personalizada

Uma **função personalizada** é uma função definida pelo usuário que é empacotada com um Excel suplemento. As funções personalizadas permitem que os desenvolvedores adicionem novas funções, além dos recursos Excel comuns, definindo essas funções em JavaScript como parte de um suplemento. Os usuários Excel podem acessar funções personalizadas da mesma forma que qualquer função nativa no Excel. Consulte [Criar funções personalizadas Excel](../excel/custom-functions-overview.md) para saber mais.

## <a name="custom-functions-runtime"></a>runtime de funções personalizadas

Um **runtime de funções personalizadas** é um runtime somente JavaScript que executa apenas funções personalizadas. Ele não tem interface do usuário e não pode interagir com Office.js APIs. Se o suplemento tiver apenas funções personalizadas, esse será um bom runtime leve a ser usado. Se suas funções personalizadas precisam interagir com o painel de tarefas ou Office.js APIs, configure um runtime de JavaScript compartilhado. Consulte [Configure seu Suplemento do Office para usar em um tempo de execução do JavaScript compartilhado](../develop/configure-your-add-in-to-use-a-shared-runtime.md) para saber mais.

Confira também: [runtime do JavaScript](#javascript-runtime), [runtime de JavaScript compartilhado, runtime compartilhado](#shared-javascript-runtime-shared-runtime).

## <a name="host"></a>host

**O host** normalmente se refere a um Office aplicativo. Os Office ou hosts que dão suporte Office suplementos são Excel, OneNote, Outlook, PowerPoint, Project e Word.

Consulte também: [aplicativo](#application), [cliente](#client), [Office aplicativo, Office cliente](#office-application-office-client).

## <a name="javascript-runtime"></a>Runtime do JavaScript

O **runtime do JavaScript** é o ambiente de host do navegador no qual o suplemento é executado. No Office no Windows e no Office no Mac, o runtime do JavaScript é um controle de navegador inserido (ou modo de exibição da Web), como Internet Explorer, Edge Legacy, Edge WebView2 ou Safari. Diferentes partes de um suplemento são executadas em runtimes separados do JavaScript. Por exemplo, comandos de suplemento, funções personalizadas e código do painel de tarefas normalmente usam runtimes javaScript separados, a menos que você configure um runtime de JavaScript compartilhado. Consulte [Navegadores usados Office suplementos para](../concepts/browsers-used-by-office-web-add-ins.md) obter mais informações.

Confira também: [runtime de funções personalizadas](#custom-functions-runtime), [runtime de JavaScript compartilhado, runtime compartilhado](#shared-javascript-runtime-shared-runtime), [modo de exibição da Web](#webview).

## <a name="office-application-office-client"></a>Office aplicativo, Office cliente

**Office cliente refere-se** a um Office aplicativo. Os Office ou clientes que dão suporte Office suplementos são Excel, OneNote, Outlook, PowerPoint, Project e Word.

Consulte também: [aplicativo](#application), [cliente](#client), [host](#host).

## <a name="platform"></a>plataforma

Uma **plataforma** geralmente se refere ao sistema operacional que executa o Office aplicativo. As plataformas que dão suporte Office suplementos incluem Windows, Mac, iPad e navegadores da Web.

## <a name="quick-start"></a>início rápido

Um **início rápido** é uma descrição de alto nível das principais habilidades e conhecimentos necessários para a operação básica de um programa específico. Na documentação Office suplementos, um início rápido é uma introdução ao desenvolvimento de um suplemento para um aplicativo específico, como Outlook. Um início rápido contém uma série de etapas que um desenvolvedor de suplementos pode concluir em aproximadamente 5 minutos, resultando em um ambiente de desenvolvimento funcional e suplemento em funcionamento.

Confira também: [tutorial](#tutorial).

## <a name="requirement-set"></a>conjunto de requisitos

[!include[Requirement set note](../includes/office-js-requirement-sets.md)]

## <a name="ribbon-ribbon-button"></a>faixa de opções, botão da faixa de opções

Uma **faixa** de opções é uma barra de comandos que organiza os recursos de um aplicativo em uma série de guias ou botões na parte superior de uma janela. Um **botão da faixa** de opções é um dos botões desta série. Consulte [Mostrar ou ocultar a faixa de opções Office](https://support.microsoft.com/office/d946b26e-0c8c-402d-a0f7-c6efa296b527#ID0EBBD=Newer_Versions) para obter mais informações.

## <a name="runtime"></a>Runtime

Consulte: [runtime do JavaScript](#javascript-runtime).

## <a name="shared-javascript-runtime-shared-runtime"></a>runtime de JavaScript compartilhado, runtime compartilhado

Um **runtime de JavaScript** compartilhado, ou **runtime** compartilhado, permite que todo o código em seu suplemento, incluindo painel de tarefas, comandos de suplemento e funções personalizadas, seja executado no mesmo runtime do JavaScript e continue em execução mesmo quando o painel de tarefas estiver fechado. Consulte Configurar seu suplemento Office para usar um [runtime e um Dicas JavaScript](../develop/configure-your-add-in-to-use-a-shared-runtime.md) compartilhados para usar o [runtime do JavaScript](https://devblogs.microsoft.com/microsoft365dev/tips-for-using-the-shared-javascript-runtime-in-your-office-add-in%e2%80%af/) compartilhado em seu suplemento do Office para saber mais.

Consulte também: [runtime de funções personalizadas](#custom-functions-runtime), [runtime do JavaScript](#javascript-runtime).

## <a name="task-pane"></a>painel de tarefas

Os painéis de tarefas são superfícies de interface ou exibições da Web que normalmente aparecem no lado direito da janela no Excel, Outlook, PowerPoint e Word. Os painéis de tarefa concedem aos usuários acesso a controles de interface que executam códigos para modificar documentos ou emails ou exibir dados de uma fonte de dados. Use painéis de tarefas quando você não precisar ou não puder inserir a funcionalidade diretamente no documento. Consulte [os painéis de tarefas Office suplementos para](../design/task-pane-add-ins.md) saber mais.

Consulte também: [modo de exibição da Web](#webview).

## <a name="tutorial"></a>Tutorial

Um **tutorial** é um auxílio de ensino projetado para ajudar as pessoas a aprender a usar um produto ou procedimento. No contexto Office suplementos, um tutorial orienta um desenvolvedor de suplementos por meio do processo completo de desenvolvimento de suplementos para um aplicativo específico, como Excel. Isso envolve seguir 20 ou mais etapas e é um investimento de tempo maior do que um [início rápido](#quick-start).

Confira também: [início rápido](#quick-start).

## <a name="custom-functions-only-add-in"></a>suplemento somente para funções personalizadas

Um suplemento que contém uma função personalizada, mas nenhuma interface do usuário, como um painel de tarefas. As funções personalizadas nesse tipo de suplemento são executadas em um runtime somente JavaScript. Uma função personalizada que inclui uma interface do usuário pode usar um runtime compartilhado ou uma combinação de um runtime somente JavaScript e um runtime de suporte a HTML. Recomendamos que, se você tiver uma interface do usuário, use um runtime compartilhado. 

Consulte também: [função personalizada](#custom-function), [runtime de funções personalizadas](#custom-functions-runtime).

## <a name="web-add-in"></a>suplemento Web

**O suplemento Web** é um termo herdado para um Office suplemento. Esse termo pode ser usado quando Microsoft 365 documentação do Microsoft 365 precisa distinguir suplementos Office modernos de outros tipos de suplementos, como VBA, COM ou VSTO.

Consulte também: [suplemento](#add-in).

## <a name="webview"></a>Webview

Um **modo de exibição** da Web é um elemento ou exibição que exibe o conteúdo da Web dentro de um aplicativo. Os suplementos de conteúdo e os painéis de tarefas contêm navegadores da Web inseridos e são exemplos de exibições da Web Office suplementos.

Consulte também: [suplemento de conteúdo](#content-add-in), [painel de tarefas](#task-pane).

## <a name="xll"></a>XLL

Um **suplemento XLL** é um Excel de suplemento que fornece funções definidas pelo usuário e tem a **extensão de arquivo .xll**. Um arquivo XLL é um tipo de arquivo DLL (biblioteca de vínculo dinâmico) que só pode ser aberto por Excel. Os arquivos de suplemento XLL devem ser gravados em C ou C++. Funções personalizadas são o equivalente moderno de funções XLL definidas pelo usuário. As funções personalizadas oferecem suporte entre plataformas e são compatíveis com versões anteriores com arquivos XLL. Consulte [Estender funções personalizadas com funções XLL definidas pelo usuário](/office/dev/add-ins/excel/make-custom-functions-compatible-with-xll-udf) para obter mais informações.

Consulte também: [função personalizada](#custom-function).

## <a name="yeoman-generator-yo-office"></a>Gerador Yeoman, seu escritório

O [gerador Yeoman para Office suplementos](../develop/yeoman-generator-overview.md) usa código aberto [ferramenta Yeoman](https://github.com/yeoman/yo) para gerar um suplemento Office por meio da linha de comando. `yo office`é o comando que executa o gerador Yeoman para Office suplementos. Os Office inícios rápidos e tutoriais de suplementos usam o gerador Yeoman.

## <a name="see-also"></a>Confira também

- [Recursos adicionais de suplementos do Office](resources-links-help.md)