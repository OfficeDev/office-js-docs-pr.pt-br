---
title: Guia do desenvolvedor do suplemento VSTO
description: Um roteiro recomendado para desenvolvedores experientes de suplemento do VSTO para recursos de aprendizagem de suplementos Web do Office.
ms.date: 10/14/2020
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 1dca15a4d286e3bfa5b7ba4a502bb9161bf3257f
ms.sourcegitcommit: 42e6cfe51d99d4f3f05a3245829d764b28c46bbb
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/23/2020
ms.locfileid: "48741061"
---
# <a name="vsto-add-in-developers-guide"></a>Guia do desenvolvedor do suplemento VSTO

Você criou alguns suplementos do VSTO para aplicativos do Office executados no Windows, e agora está aprendendo um nova maneira de estender o Office que será executado no Windows, no Mac e na versão online do pacote do Office: suplementos Web do Office.

Sua compreensão sobre os modelos de objeto para Excel, Word e outros aplicativos do Office será uma grande ajuda, pois os modelos de objeto nos suplementos Web do Office seguem padrões semelhantes. Mas haverá alguns desafios:

- Você trabalhará com uma linguagem diferente (JavaScript ou TypeScript) em vez de C# ou Visual Basic .NET. (Há também uma maneira, descrita abaixo, de reutilizar alguns de seus códigos existentes em um suplemento Web).
- Os suplementos Web do Office são implantados de forma diferente dos suplementos do VSTO.
- Os suplementos Web do Office são aplicativos Web executados em uma janela simplificada do navegador que está incorporada ao aplicativo do Office. Portanto, é necessário obter um conhecimento básico dos aplicativos Web e de como eles são hospedados em servidores Web ou em contas de nuvem. 

Por esses motivos, uma boa parte deste artigo duplica o caminho de aprendizagem para iniciantes das extensões do Office: [Guia para iniciantes](learning-path-beginner.md). O que adicionamos são alguns recursos de aprendizagem adicionais para ajudar os desenvolvedores do suplemento VSTO a aproveitar suas experiências e também a reutilizarem o código existente.

## <a name="step-0-prerequisites"></a>Etapa 0: Pré-requisitos

- Os suplementos Web do Office (também chamados de suplementos do Office) são essencialmente aplicativos Web incorporados no Office. Portanto, você deve primeiro ter um conhecimento básico dos aplicativos Web e de como eles são hospedados na Web. Há uma quantidade enorme de informações sobre isso na Internet, em livros e em cursos online. Uma boa maneira de começar, se você não tem nenhum conhecimento prévio sobre aplicativos da Web, é procurar "O que é um aplicativo da Web?" no Bing.
- A principal linguagem de programação que você usará na criação de suplementos do Office é o JavaScript ou o TypeScript. Pense no TypeScript como uma versão fortemente tipada do JavaScript. Se você não conhece nenhuma dessas linguagens, mas tem experiência com VBA, VB.Net e C#, provavelmente achará o TypeScript mais fácil de aprender. Novamente, há muitas informações sobre essas linguagens de programação na Internet, em livros e em cursos online.

## <a name="step-1-begin-with-fundamentals"></a>Etapa 1: Comece com os fundamentos

Sabemos que você está ansioso para começar a codificar, mas há algumas coisas sobre os Suplementos do Office que você deve ler antes de abrir o IDE ou o editor de código.

- [Visão Geral da Plataforma de Suplementos do Office](office-add-ins.md): Descubra o que são os suplementos da Web do Office e como eles diferem das formas mais antigas de estender o Office, como os suplementos do VSTO.
- [Desenvolva Suplementos do Office](../develop/develop-overview.md): Obtenha uma visão geral do desenvolvimento e do ciclo de vida do Suplemento do Office, incluindo ferramentas, criando uma interface de usuário do suplemento e usando as APIs de JavaScript para interagir com o documento do Office.

Existem muitos links nesses artigos, mas se você estiver migrando para os suplementos Web do Office, recomendamos que você volte aqui quando os tiver lido e continue na próxima seção.

## <a name="step-2-install-tools-and-create-your-first-add-in"></a>Etapa 2: Instale ferramentas e crie o seu primeiro suplemento

Agora você tem uma visão geral, então comece com um de nossos inícios rápidos. Para fins de aprendizado da plataforma, recomendamos o início rápido do Excel. Há uma versão baseada no Visual Studio e outra baseada em Node.js e Visual Studio Code. Se você estiver migrando de suplementos do VSTO, provavelmente encontrará a versão do Visual Studio mais fácil de trabalhar.

- [Visual Studio](../quickstarts/excel-quickstart-jquery.md?tabs=visualstudio)
- [Node.js e Visual Studio Code](../quickstarts/excel-quickstart-jquery.md?tabs=yeomangenerator)

## <a name="step-3-code"></a>Etapa 3: Codifique

Não se pode aprender a dirigir lendo o manual do proprietário, então comece a codificar com este [tutorial do Excel](../tutorials/excel-tutorial.md). Você usará a biblioteca JavaScript do Office e um pouco de XML no manifesto dos suplementos. Não é necessário memorizar nada, porque você terá mais informações sobre ambos em etapas posteriores.

## <a name="step-4-understand-the-javascript-library"></a>Etapa 4: Entenda a biblioteca JavaScript

Obtenha uma visão geral da biblioteca JavaScript do Office com este tutorial do Microsoft Learn: [Entenda as APIs JavaScript do Office](/learn/modules/intro-office-add-ins/3-apis).

Em seguida, explore as APIs do Office JavaScript com a [ferramenta Script Lab](explore-with-script-lab.md) – uma área restrita para executar e explorar as APIs.

### <a name="special-resource-for-vsto-add-in-developers"></a>Um recurso especial para desenvolvedores de suplemento do VSTO

Esse seria um bom lugar para dar uma olhada no exemplo de suplemento, [Suplemento do Excel JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker). Ele foi criado para destacar as semelhanças e diferenças entre suplementos do VSTO e suplementos Web do Office, e o leiame do exemplo indica os pontos importantes da comparação.

## <a name="step-5-understand-the-manifest"></a>Etapa 5: Entenda o manifesto

Entenda os objetivos do manifesto de suplemento Web e veja uma introdução à sua marcação XML no [Manifesto XML dos suplementos do Office](../develop/add-in-manifests.md).

## <a name="step-6-for-vsto-developers-only-reuse-your-vsto-code"></a>Etapa 6 (somente para desenvolvedores do VSTO): Reutilize seu código de VSTO

Você pode reutilizar alguns dos códigos de suplemento do VSTO em um suplemento Web do Office, movendo-os para o back-end do seu aplicativo Web no servidor e disponibilizando-o para o JavaScript ou TypeScript como uma API da Web. Para obter instruções, confira [Tutorial: compartilhar código entre um suplemento do VSTO e um suplemento do Office usando uma biblioteca de códigos compartilhados](../tutorials/migrate-vsto-to-office-add-in-shared-code-library-tutorial.md).

## <a name="next-steps"></a>Próximas etapas

Parabéns por concluir o roteiro de aprendizagem para desenvolvedores de suplementos VSTO para suplementos Web do Office! Veja algumas sugestões para explorar ainda mais a documentação:

- Tutoriais ou inícios rápidos para outros aplicativos do Office:

  - [Início rápido do OneNote](../quickstarts/onenote-quickstart.md)
  - [Tutorial do Outlook](/outlook/add-ins/addin-tutorial)
  - [Tutorial do PowerPoint](../tutorials/powerpoint-tutorial.md)
  - [Início rápido do Project](../quickstarts/project-quickstart.md)
  - [Tutorial do Word](../tutorials/word-tutorial.md)

- Outros assuntos importantes:

  - [Desenvolver Suplementos do Office ](../develop/develop-overview.md)
  - [Práticas recomendadas para o desenvolvimento de suplementos do Office](../concepts/add-in-development-best-practices.md)
  - [Fazer o design de Suplementos do Office](../design/add-in-design.md)
  - [Testar e depurar Suplementos do Office](../testing/test-debug-office-add-ins.md)
  - [Implantar e publicar Suplementos do Office](../publish/publish.md)
  - [Recursos](../resources/resources-links-help.md)
  - [Saiba mais sobre o Programa para Desenvolvedores do Microsoft 365](https://developer.microsoft.com/microsoft-365/dev-program)
