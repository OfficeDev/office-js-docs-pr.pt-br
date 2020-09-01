---
title: Desenvolver Suplementos do Office com o Visual Studio
description: Como desenvolver Suplementos do Office com o Visual Studio
ms.date: 12/31/2019
localization_priority: Priority
ms.openlocfilehash: ae627b09b9160abc01deec6d52abeb922f02c833
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292824"
---
# <a name="develop-office-add-ins-with-visual-studio"></a>Desenvolver Suplementos do Office com o Visual Studio

Esse artigo descreve como usar o Visual Studio para desenvolver um suplemento do Office. Caso você já tenha criado seu suplemento, você pode pular para a seção [Desenvolver o suplemento usando o Visual Studio](#develop-the-add-in-using-visual-studio).

> [!NOTE]
> Como uma alternativa ao uso do Visual Studio, você pode optar por usar o gerador Yeoman para suplementos do Office e o VS Code para criar um suplemento do Office. Para obter mais informações sobre essa opção, confira [Criar um suplemento do Office](../overview/office-add-ins-fundamentals.md#creating-an-office-add-in).

## <a name="create-the-add-in-project-using-visual-studio"></a>Criar o projeto de suplemento usando o Visual Studio

O Visual Studio pode ser usado para criar suplementos do Office para o Excel, Outlook, Word e PowerPoint. Um projeto do suplemento do Office é criado como parte de uma solução do Visual Studio e usa HTML, CSS e JavaScript. Para criar um suplemento do Office com o Visual Studio, siga as instruções no início rápido que correspondam ao suplemento que você deseja criar:

- [Início rápido do Excel](../quickstarts/excel-quickstart-jquery.md?tabs=visualstudio)
- [Início rápido do Outlook](../quickstarts/outlook-quickstart.md?tabs=visualstudio)
- [Início rápido do Word](../quickstarts/word-quickstart.md?tabs=visualstudio)
- [Início rápido do PowerPoint](../quickstarts/powerpoint-quickstart.md?tabs=visualstudio)

O Visual Studio não tem suporte para a criação de suplementos do Office para o OneNote ou Project. Para criar suplementos do Office para qualquer um desses aplicativos, você precisará usar o gerador Yeoman para Suplementos do Office, conforme descrito no [Início rápido do OneNote](../quickstarts/onenote-quickstart.md) ou no [Início rápido do Project](../quickstarts/project-quickstart.md).

## <a name="develop-the-add-in-using-visual-studio"></a>Desenvolver o suplemento usando o Visual Studio

O Visual Studio cria um suplemento básico com funcionalidade limitada. Você pode personalizar o suplemento editando o [manifesto](add-in-manifests.md), HTML, JavaScript e arquivos CSS no Visual Studio. Para obter uma descrição de alto nível da estrutura do projeto e dos arquivos no projeto de suplemento que o Visual Studio cria, confira a orientação do Visual Studio no início rápido concluído para criar seu suplemento. 

> [!TIP]
> Como um suplemento do Office é um aplicativo da Web, você precisará de pelo menos habilidades básicas de desenvolvimento na Web para personalizar seu suplemento. Se você não conhece o JavaScript, recomendamos que revise o [tutorial do Mozilla JavaScript](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).

Para personalizar o seu suplemento, você precisará compreender os conceitos descritos na área [Conceitos básicos > Desenvolver](develop-overview.md) dessa documentação, além dos conceitos descritos na área de documentação específica do aplicativo que corresponde ao suplemento que você está criando (por exemplo, o [Excel](../excel/index.yml)). 

## <a name="test-and-debug-the-add-in"></a>Testar e depurar o suplemento

Os métodos para testar, depurar e solucionar problemas de Suplementos do Office variam de acordo com a plataforma. Para saber mais, confira [Depurar Suplementos do Office no Visual Studio](debug-office-add-ins-in-visual-studio.md) e [Testar e depurar Suplementos do Office](../testing/test-debug-office-add-ins.md).

## <a name="publish-the-add-in"></a>Publique o suplemento

Um Suplemento do Office é formado por um aplicativo Web e um arquivo de manifesto. O aplicativo Web define a interface do usuário e a funcionalidade do suplemento, enquanto o manifesto especifica o local do aplicativo Web e define as configurações e os recursos do suplemento.

Enquanto você estiver desenvolvendo seu suplemento no Visual Studio, seu suplemento será executado no seu servidor Web local (`localhost`). Quando o suplemento estiver funcionando como desejado e você estiver pronto para publicá-lo para que outros usuários acessem-no, será necessário concluir as seguintes etapas:

1. Implantar o aplicativo Web em um servidor Web ou serviço de hospedagem na Web (por exemplo, Microsoft Azure).
2. Atualize o manifesto para especificar a URL do aplicativo implantado. 
3. Escolha o método que deseja usar para [implantar seu suplemento do Office](../publish/publish.md) e siga as instruções para publicar o arquivo de manifesto.

## <a name="see-also"></a>Confira também

- [Criando Suplementos do Office ](../overview/office-add-ins-fundamentals.md)
- [Principais conceitos dos Suplementos do Office](../overview/core-concepts-office-add-ins.md)
- [Desenvolver Suplementos do Office](../develop/develop-overview.md)
- [Fazer o design de Suplementos do Office](../design/add-in-design.md)
- [Testar e depurar Suplementos do Office](../testing/test-debug-office-add-ins.md)
- [Publicar Suplementos do Office](../publish/publish.md)
