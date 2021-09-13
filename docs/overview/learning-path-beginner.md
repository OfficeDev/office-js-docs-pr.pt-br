---
title: Guia para iniciantes em suplementos do Office
description: Um caminho recomendado para iniciantes, através dos recursos de aprendizado dos Suplementos do Office.
ms.date: 02/12/2021
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: beafa972129133955ddebef9ebaaeae6dda101b6
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59151768"
---
# <a name="beginners-guide"></a>Guia do iniciante

Deseja começar a criar suas próprias extensões do Office para várias plataformas? As etapas a seguir mostram o que ler primeiro, quais ferramentas instalar e os tutoriais recomendados a serem concluídos.

> [!NOTE]
> Se você não tiver experiência em criar suplementos VSTO do Office, recomendamos que você acesse imediatamente o [guia do desenvolvedor do suplemento VSTO](learning-path-transition.md), que é um superconjunto das informações deste artigo.

## <a name="step-0-prerequisites"></a>Etapa 0: Pré-requisitos

- Os suplementos do Office são essencialmente aplicativos da Web incorporados ao Office. Portanto, você deve primeiro ter um entendimento básico dos aplicativos da Web e de como eles são hospedados na Web. Há uma quantidade enorme de informações sobre isso na Internet, em livros e em cursos online. Uma boa maneira de começar, se você não tem nenhum conhecimento prévio sobre aplicativos da Web, é procurar "O que é um aplicativo da Web?" no Bing.
- A principal linguagem de programação que você usará na criação de Suplementos do Office é JavaScript ou TypeScript. Pense no TypeScript como uma versão fortemente tipada do JavaScript. Se você não conhece nenhuma dessas linguagens de programação, mas possui experiência com VBA, VB.Net e C#, provavelmente achará o TypeScript mais fácil de aprender. Novamente, há muitas informações sobre essas linguagens de programação na Internet, em livros e em cursos online.

## <a name="step-1-begin-with-fundamentals"></a>Etapa 1: Comece com os fundamentos

Sabemos que você está ansioso para começar a codificar, mas há algumas coisas sobre os Suplementos do Office que você deve ler antes de abrir o IDE ou o editor de código.

- [Visão Geral da Plataforma de Suplementos do Office](office-add-ins.md): Descubra o que são os suplementos da Web do Office e como eles diferem das formas mais antigas de estender o Office, como os suplementos do VSTO.
- [Desenvolva Suplementos do Office](../develop/develop-overview.md): Obtenha uma visão geral do desenvolvimento e ciclo de vida do Suplemento do Office, incluindo as ferramentas, criando um suplemento na Interface do Usuário e usando APIs de JavaScript para interagir com o documento do Office.

Existem muitos links nesses artigos, mas se você é iniciante nos suplementos do Office, recomendamos que você volte aqui quando os tiver lido e continue na próxima seção.

## <a name="step-2-install-tools-and-create-your-first-add-in"></a>Etapa 2: Instale ferramentas e crie o seu primeiro suplemento

Agora você tem uma visão geral, então comece com um de nossos inícios rápidos. Para fins de aprendizado da plataforma, recomendamos o início rápido do Excel. Há uma versão baseada no Visual Studio e uma versão baseada no Node.js e no Visual Studio Code.

- [Visual Studio](../quickstarts/excel-quickstart-jquery.md?tabs=visualstudio)
- [Node.js e Visual Studio Code](../quickstarts/excel-quickstart-jquery.md?tabs=yeomangenerator)

## <a name="step-3-code"></a>Etapa 3: Codifique

Não se pode aprender a dirigir lendo o manual do proprietário, então comece a codificar com este [tutorial do Excel](../tutorials/excel-tutorial.md). Você usará a biblioteca JavaScript do Office e um pouco de XML no manifesto dos suplementos. Não é necessário memorizar nada, porque você terá mais informações sobre ambos em etapas posteriores.

## <a name="step-4-understand-the-javascript-library"></a>Etapa 4: Entenda a biblioteca JavaScript

Primeiro, obtenha uma visão geral da biblioteca JavaScript do Office com este tutorial do Microsoft Learn: [Entenda as APIs JavaScript do Office](/learn/modules/understand-office-javascript-apis/index).

Em seguida, explore as APIs do Office JavaScript com a [ferramenta Script Lab](explore-with-script-lab.md) -- uma área restrita para executar e explorar as APIs.

## <a name="step-5-understand-the-manifest"></a>Etapa 5: Entenda o manifesto

Entenda os objetivos do manifesto do suplemento e veja uma introdução à sua marcação XML no [Manifesto XML dos Suplementos do Office](../develop/add-in-manifests.md).

## <a name="next-steps"></a>Próximas Etapas

Parabéns por concluir o caminho de aprendizado para iniciantes dos Suplementos do Office! Veja algumas sugestões para explorar ainda mais a documentação:

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
