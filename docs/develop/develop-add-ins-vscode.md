---
title: Desenvolver Suplementos do Office com o Código do Visual Studio
description: Como desenvolver Suplementos do Office com o Visual Studio Code.
ms.date: 02/18/2022
ms.localizationpriority: high
ms.openlocfilehash: 6710884a9bc751e6a94607581223dabaea0bce3b
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/08/2022
ms.locfileid: "63511291"
---
# <a name="develop-office-add-ins-with-visual-studio-code"></a>Desenvolver Suplementos do Office com o Código do Visual Studio

Este artigo descreve como usar [o Código do Visual Studio (VS Code)](https://code.visualstudio.com) para desenvolver um suplemento do Office.

> [!NOTE]
> Para saber mais sobre como usar o Visual Studio para criar um suplemento do Office, confira [Desenvolver suplementos do Office com o Visual Studio](develop-add-ins-visual-studio.md).

## <a name="prerequisites"></a>Pré-requisitos

- [Código do Visual Studio](https://code.visualstudio.com/)

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-the-add-in-project-using-the-yeoman-generator"></a>Criar um projeto de suplemento usando o gerador Yeoman

Se você estiver usando o VS Code como o seu ambiente de desenvolvimento integrado (IDE), crie o projeto do Suplemento do Office com o [Gerador Yeoman para Suplementos do Office](https://github.com/OfficeDev/generator-office). O gerador Yeoman cria um projeto Node.js que pode ser gerenciado com o VS Code ou qualquer outro editor.

Para criar um Suplemento do Office com o gerador Yeoman, siga as instruções em[início rápido em 5 minutos](../index.yml) que corresponda ao tipo de suplemento que você deseja criar.

## <a name="develop-the-add-in-using-vs-code"></a>Desenvolver o suplemento usando o VS Code

Quando o gerador Yeoman terminar de criar o projeto do suplemento, abra a pasta raiz do projeto com o VS Code.

[!INCLUDE [Instructions for opening add-in project in VS Code via command line](../includes/vs-code-open-project-via-command-line.md)]

O gerador Yeoman cria um suplemento básico com funcionalidade limitada. Você pode personalizar o suplemento editando o [manifesto](add-in-manifests.md), HTML, JavaScript ou TypeScript e arquivos CSS no VS Code. Para obter uma descrição de alto nível sobre a estrutura e os arquivos do projeto no projeto de suplemento que o gerador de Yeoman cria, confira o tópico diretrizes do gerador Yeoman dentro em [Início rápido em 5 minutos](../index.yml) que corresponda ao tipo de suplemento que você criou.

## <a name="test-and-debug-the-add-in"></a>Testar e depurar o suplemento

Os métodos para testar, depurar e solucionar problemas de Suplementos do Office variam de acordo com a plataforma. Para mais informações, confira [Testar e depurar Suplementos do Office](../testing/test-debug-office-add-ins.md).

## <a name="publish-the-add-in"></a>Publique o suplemento

[!include[instructions for publishing an Office Add-in](../includes/publish-add-in.md)]

## <a name="see-also"></a>Confira também

- [Principais conceitos dos Suplementos do Office](../overview/core-concepts-office-add-ins.md)
- [Desenvolver Suplementos do Office ](../develop/develop-overview.md)
- [Fazer o design de Suplementos do Office](../design/add-in-design.md)
- [Testar e depurar Suplementos do Office](../testing/test-debug-office-add-ins.md)
- [Publicar Suplementos do Office](../publish/publish.md)