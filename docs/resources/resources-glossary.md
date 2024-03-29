---
title: Glossário de termos de Suplementos do Office
description: Um glossário de termos comumente usados em toda a documentação de Suplementos do Office.
ms.date: 09/14/2022
ms.localizationpriority: medium
ms.openlocfilehash: ef8df6e344698f7d67ebe7afe1759e13630b385d
ms.sourcegitcommit: cff5d3450f0c02814c1436f94cd1fc1537094051
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/30/2022
ms.locfileid: "68234911"
---
# <a name="office-add-ins-glossary"></a>Glossário de Suplementos do Office

Este é um glossário de termos comumente usados em toda a documentação de Suplementos do Office.

## <a name="add-in"></a>suplemento

Os Suplementos do Office são aplicativos Web que estendem aplicativos do Office. Esses aplicativos Web adicionam novas funcionalidades ao aplicativo do Office, como trazer dados externos, automatizar processos ou inserir objetos interativos em documentos do Office.

Os Suplementos do Office diferem dos suplementos VBA, COM e VSTO porque oferecem suporte multiplataforma (geralmente Web, Windows, Mac e iPad) e são baseados em tecnologias web padrão (HTML, CSS e JavaScript). A linguagem de programação principal de um suplemento do Office é JavaScript ou TypeScript.

## <a name="add-in-commands"></a>comandos de suplemento

**Comandos de suplemento são** elementos de interface do usuário, como botões e menus, que estendem a interface do usuário do Office para seu suplemento. Quando os usuários selecionam um elemento de comando de suplemento, eles iniciam ações como executar código JavaScript ou exibir o suplemento em um painel de tarefas. Os comandos de suplemento permitem que o suplemento se sinta como parte do Office, o que dá aos usuários mais confiança em seu suplemento. Consulte [comandos de suplemento para comandos do Excel, PowerPoint e Word](../design/add-in-commands.md) e [Suplementos para o Outlook](../outlook/add-in-commands-for-outlook.md) para saber mais.

Consulte também: faixa [de opções, botão da faixa de opções](#ribbon-ribbon-button).

## <a name="application"></a>aplicação

**O** aplicativo refere-se a um aplicativo do Office. Os aplicativos do Office que dão suporte a Suplementos do Office são Excel, OneNote, Outlook, PowerPoint, Project e Word.

Consulte também: [cliente](#client), [host](#host), [aplicativo do Office, cliente do Office](#office-application-office-client).

## <a name="application-specific-api"></a>API específica do aplicativo

As APIs específicas do aplicativo fornecem objetos fortemente tipados que interagem com objetos nativos de um aplicativo específico do Office. Por exemplo, você chama as APIs JavaScript do Excel para acesso a planilhas, intervalos, tabelas, gráficos e muito mais. As APIs específicas do aplicativo estão disponíveis no momento para Excel, OneNote, PowerPoint, Visio e Word. Consulte [o modelo de API específico do aplicativo](../develop/application-specific-api-model.md) para saber mais.

Consulte também: [API comum](#common-api).

## <a name="client"></a>Cliente

**O** cliente normalmente se refere a um aplicativo do Office. Os aplicativos do Office ou clientes que dão suporte a Suplementos do Office são Excel, OneNote, Outlook, PowerPoint, Project e Word.

Consulte também: [aplicativo](#application), [host](#host), [aplicativo do Office, cliente do Office](#office-application-office-client).

## <a name="common-api"></a>Common API

APIs comuns são usadas para acessar recursos como interface do usuário, caixas de diálogo e configurações de cliente que são comuns em vários aplicativos do Office. Esse modelo de API usa [retornos de chamada](https://developer.mozilla.org/docs/Glossary/Callback_function), que permitem especificar apenas uma operação em cada solicitação enviada ao aplicativo do Office.

As APIs comuns foram introduzidas com o Office 2013 e são usadas para interagir com o Office 2013 ou posterior. Algumas APIs comuns são APIs herdadas do início de 2010. Excel, PowerPoint e Word têm funcionalidade de API comum, mas a maior parte dessa funcionalidade foi substituída ou substituída pelo modelo de API específico do aplicativo. As APIs específicas do aplicativo são preferenciais quando possível.

Outras APIs comuns, como as APIs comuns relacionadas ao Outlook, à interface do usuário e à autenticação, são as APIs modernas e preferenciais para essas finalidades. Para obter detalhes sobre o modelo de objeto da API Comum, consulte [o modelo de objeto da API JavaScript comum](../develop/office-javascript-api-object-model.md).

Consulte também: [API específica do aplicativo](#application-specific-api).

## <a name="content-add-in"></a>suplemento de conteúdo

**Os suplementos de conteúdo são** modos de exibição da Web ou exibições do navegador da Web que são inseridos diretamente em documentos do Excel, OneNote ou PowerPoint. Os suplementos de conteúdo concedem aos usuários acesso a controles de interface que executam códigos para modificar documentos ou exibir dados de uma fonte de dados. Use suplementos de conteúdo quando quiser inserir a funcionalidade diretamente no documento. Consulte [Suplementos do Office de Conteúdo](../design/content-add-ins.md) para saber mais.

Consulte também: [modo de exibição da Web](#webview).

## <a name="content-delivery-network-cdn"></a>CDN (rede de distribuição de conteúdo)

Uma **rede de distribuição de** conteúdo **ou CDN** é uma rede distribuída de servidores e data centers. Normalmente, ele fornece maior disponibilidade e desempenho de recursos quando comparado a um único servidor ou data center.

## <a name="contoso"></a>Contoso

**Contoso** Ltd. (também conhecida como Contoso e Contoso University) é uma empresa fictícia usada pela Microsoft como uma empresa e domínio de exemplo.

## <a name="custom-function"></a>função personalizada

Uma **função personalizada** é uma função definida pelo usuário que é empacotada com um suplemento do Excel. As funções personalizadas permitem que os desenvolvedores adicionem novas funções, além dos recursos típicos do Excel, definindo essas funções em JavaScript como parte de um suplemento. Os usuários no Excel podem acessar funções personalizadas da mesma forma que qualquer função nativa no Excel. Consulte [Criar funções personalizadas no Excel](../excel/custom-functions-overview.md) para saber mais.

## <a name="custom-functions-runtime"></a>runtime de funções personalizadas

Um **runtime de funções personalizadas** é um [runtime somente JavaScript](../testing/runtimes.md#javascript-only-runtime) que executa funções personalizadas em algumas combinações de host e plataforma do Office. Ele não tem interface do usuário e não pode interagir com Office.js APIs. Se o suplemento tiver apenas funções personalizadas, esse será um bom runtime leve a ser usado. Se suas funções personalizadas precisam interagir com o painel de tarefas ou Office.js APIs, configure um [runtime compartilhado](../testing/runtimes.md#shared-runtime). Consulte [Configurar seu Suplemento do Office para usar um runtime compartilhado](../develop/configure-your-add-in-to-use-a-shared-runtime.md) para saber mais.

Consulte também: [runtime](#runtime), [runtime compartilhado](#shared-runtime).

## <a name="custom-functions-only-add-in"></a>suplemento somente para funções personalizadas

Um suplemento que contém uma função personalizada, mas nenhuma interface do usuário, como um painel de tarefas. As funções personalizadas nesse tipo de suplemento são executadas em um [runtime somente JavaScript](../testing/runtimes.md#javascript-only-runtime). Uma função personalizada que inclui uma interface do usuário pode usar um runtime compartilhado ou uma combinação de um runtime somente JavaScript e um runtime de suporte a HTML. Recomendamos que, se você tiver uma interface do usuário, use um runtime compartilhado.

Consulte também: [função personalizada](#custom-function), [runtime de funções personalizadas](#custom-functions-runtime).

## <a name="host"></a>host

**\<Host\>** normalmente se refere a um aplicativo do Office. Os aplicativos do Office ou hosts que dão suporte a Suplementos do Office são Excel, OneNote, Outlook, PowerPoint, Project e Word.

Consulte também: [aplicativo](#application), [cliente](#client), aplicativo [do Office, cliente do Office](#office-application-office-client).

## <a name="office-application-office-client"></a>Aplicativo do Office, cliente do Office

**O cliente do Office** refere-se a um aplicativo do Office. Os aplicativos do Office ou clientes que dão suporte a Suplementos do Office são Excel, OneNote, Outlook, PowerPoint, Project e Word.

Consulte também: [aplicativo](#application), [cliente](#client), [host](#host).

## <a name="perpetual"></a>Perpétuo

**Perpétuo** refere-se a versões do Office disponíveis por meio de um contrato de licenciamento por volume ou canais de varejo.

Outro conteúdo da Microsoft pode usar o **termo não assinatura** para representar esse conceito.

Consulte também: [varejo, perpétuo](#retail-retail-perpetual) de varejo, [licenciado por volume, licença por volume perpétua, licenciamento por volume](#volume-licensed-volume-licensed-perpetual-volume-licensing)

## <a name="platform"></a>plataforma

Uma **plataforma** geralmente se refere ao sistema operacional que executa o aplicativo do Office. As plataformas que dão suporte a Suplementos do Office incluem navegadores da Web, Windows, Mac, iPad e Windows.

## <a name="quick-start"></a>início rápido

Um **início rápido** é uma descrição de alto nível das principais habilidades e conhecimentos necessários para a operação básica de um programa específico. Na documentação de Suplementos do Office, um início rápido é uma introdução ao desenvolvimento de um suplemento para um aplicativo específico, como o Outlook. Um início rápido contém uma série de etapas que um desenvolvedor de suplementos pode concluir em aproximadamente 5 minutos, resultando em um ambiente de desenvolvimento funcional e suplemento em funcionamento.

Confira também: [tutorial](#tutorial).

## <a name="requirement-set"></a>conjunto de requisitos

[!include[Requirement set note](../includes/office-js-requirement-sets.md)]

## <a name="retail-retail-perpetual"></a>varejo, varejo perpétuo

**Varejo** refere-se a versões perpétuas do Office disponíveis por meio de canais de varejo. Elas não incluem versões fornecidas por uma assinatura do Microsoft 365 nem por um contrato de licenciamento por volume.

Outro conteúdo da Microsoft pode usar o termo compra **única** ou **consumidor** para representar esse conceito.

Consulte também: [perpétuo](#perpetual)

## <a name="ribbon-ribbon-button"></a>faixa de opções, botão da faixa de opções

Uma **faixa** de opções é uma barra de comandos que organiza os recursos de um aplicativo em uma série de guias ou botões na parte superior de uma janela. Um **botão da faixa** de opções é um dos botões desta série. Consulte [Mostrar ou ocultar a faixa de opções no Office](https://support.microsoft.com/office/d946b26e-0c8c-402d-a0f7-c6efa296b527#ID0EBBD=Newer_Versions) para obter mais informações.

## <a name="runtime"></a>Runtime

Um **runtime** é o ambiente de host (incluindo um mecanismo JavaScript e geralmente também um mecanismo de renderização HTML) no qual o suplemento é executado. No Office no Windows e no Office no Mac, o runtime é um controle de navegador inserido (ou modo de exibição da Web), como Internet Explorer, Edge Legacy, Edge WebView2 ou Safari. Diferentes partes de um suplemento são executadas em runtimes separados. Por exemplo, comandos de suplemento, funções personalizadas e código do painel de tarefas normalmente usam runtimes separados, a menos que você configure um [runtime compartilhado](../testing/runtimes.md#shared-runtime). Consulte [Runtimes em Suplementos e](../testing/runtimes.md) [Navegadores do Office usados pelos Suplementos do Office](../concepts/browsers-used-by-office-web-add-ins.md) para obter mais informações.

Consulte também: [runtime de funções personalizadas](#custom-functions-runtime), [runtime compartilhado](#shared-runtime), [modo de exibição da Web](#webview).

## <a name="shared-runtime"></a>runtime compartilhado

Um **runtime** compartilhado permite que todo o código no suplemento, incluindo painel de tarefas, comandos de suplemento e funções personalizadas, seja executado no mesmo runtime e continue em execução mesmo quando o painel de tarefas estiver fechado. Confira [o runtime compartilhado](../testing/runtimes.md#shared-runtime) e [dicas para usar o runtime compartilhado em seu Suplemento do Office](https://devblogs.microsoft.com/microsoft365dev/tips-for-using-the-shared-javascript-runtime-in-your-office-add-in%e2%80%af/) para saber mais.

Consulte também: [runtime de funções personalizadas](#custom-functions-runtime), [runtime](#runtime).

## <a name="subscription"></a>Assinatura

**A** assinatura refere-se a versões do Office disponíveis com uma assinatura do Microsoft 365.

## <a name="task-pane"></a>painel de tarefas

Os painéis de tarefas são superfícies de interface ou exibições da Web que normalmente aparecem no lado direito da janela no Excel, Outlook, PowerPoint e Word. Os painéis de tarefa concedem aos usuários acesso a controles de interface que executam códigos para modificar documentos ou emails ou exibir dados de uma fonte de dados. Use painéis de tarefas quando você não precisar ou não puder inserir a funcionalidade diretamente no documento. Consulte [painéis de tarefas em Suplementos do Office](../design/task-pane-add-ins.md) para saber mais.

Consulte também: [modo de exibição da Web](#webview).

## <a name="tutorial"></a>Tutorial

Um **tutorial** é um auxílio de ensino projetado para ajudar as pessoas a aprender a usar um produto ou procedimento. No contexto de Suplementos do Office, um tutorial orienta um desenvolvedor de suplementos pelo processo completo de desenvolvimento de suplementos para um aplicativo específico, como o Excel. Isso envolve seguir 20 ou mais etapas e é um investimento de tempo maior do que um [início rápido](#quick-start).

Confira também: [início rápido](#quick-start).

## <a name="volume-licensed-volume-licensed-perpetual-volume-licensing"></a>licenciamento perpétuo, perpétuo licenciado por volume e licenciado por volume

**Licenciado por volume** refere-se a uma versão perpétua do Office disponível por meio de um contrato de licenciamento por volume entre a Microsoft e sua empresa.

Outro conteúdo da Microsoft pode usar o termo **comercial** para representar esse conceito.

Consulte também: [perpétuo](#perpetual)

## <a name="web-add-in"></a>suplemento Web

**O suplemento Web** é um termo herdado para um Suplemento do Office. Esse termo pode ser usado quando a documentação do Microsoft 365 precisa distinguir suplementos modernos do Office de outros tipos de suplementos, como VBA, COM ou VSTO.

Consulte também: [suplemento](#add-in).

## <a name="webview"></a>Webview

Um **modo de exibição** da Web é um elemento ou exibição que exibe o conteúdo da Web dentro de um aplicativo. Os suplementos de conteúdo e os painéis de tarefas contêm navegadores da Web inseridos e são exemplos de exibições da Web em Suplementos do Office.

Consulte também: [suplemento de conteúdo](#content-add-in), [painel de tarefas](#task-pane).

## <a name="xll"></a>XLL

Um **suplemento XLL** é um arquivo de suplemento do Excel que fornece funções definidas pelo usuário e tem a **extensão de arquivo .xll**. Um arquivo XLL é um tipo de arquivo DLL (biblioteca de vínculo dinâmico) que só pode ser aberto pelo Excel. Os arquivos de suplemento XLL devem ser gravados em C ou C++. Funções personalizadas são o equivalente moderno de funções XLL definidas pelo usuário. As funções personalizadas oferecem suporte entre plataformas e são compatíveis com versões anteriores com arquivos XLL. Consulte [Estender funções personalizadas com funções XLL definidas pelo usuário](/office/dev/add-ins/excel/make-custom-functions-compatible-with-xll-udf) para obter mais informações.

Consulte também: [função personalizada](#custom-function).

## <a name="yeoman-generator-yo-office"></a>Gerador Yeoman, seu escritório

O [gerador Yeoman para Suplementos do Office](../develop/yeoman-generator-overview.md) usa a ferramenta [código aberto Yeoman](https://github.com/yeoman/yo) para gerar um Suplemento do Office por meio da linha de comando. `yo office` é o comando que executa o gerador Yeoman para Suplementos do Office. Os tutoriais e inícios rápidos dos Suplementos do Office usam o gerador Yeoman.

## <a name="see-also"></a>Confira também

- [Recursos adicionais de suplementos do Office](resources-links-help.md)