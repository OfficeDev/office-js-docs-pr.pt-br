---
title: Definir seu ambiente de desenvolvimento
description: Configurar seu ambiente de desenvolvedor para criar suplementos do Office
ms.date: 04/03/2020
localization_priority: Normal
ms.openlocfilehash: 03cf525408123df9e8356555f2ab7732ed2ca263
ms.sourcegitcommit: c3bfea0818af1f01e71a1feff707fb2456a69488
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/08/2020
ms.locfileid: "43185600"
---
# <a name="set-up-your-development-environment"></a>Definir seu ambiente de desenvolvimento

Este guia ajuda você a configurar ferramentas para que você possa criar suplementos do Office seguindo nosso início rápido ou tutoriais. Você precisará instalar as ferramentas na lista abaixo. Se você já tiver estes instalados, você está pronto para iniciar um início rápido, como este [início rápido reagir do Excel](../quickstarts/excel-quickstart-react.md).

- Node.js
- npm
- Uma conta do Office 365 (a versão de assinatura do Office)
- Um editor de código de sua escolha

Este guia pressupõe que você saiba como usar uma ferramenta de linha de comando. 

## <a name="install-nodejs"></a>Instalar o Node. js

Node. js é um tempo de execução de JavaScript, você precisará desenvolver suplementos do Office modernos.

Instale o Node. js [baixando a versão mais recente recomendada do site](https://nodejs.org). Siga as instruções de instalação do seu sistema operacional.

## <a name="install-npm"></a>Instalar o NPM

o NPM é um registro de software de código aberto do qual baixar os pacotes usados no desenvolvimento de suplementos do Office.

Para instalar o NPM, execute o seguinte na linha de comando:

```command&nbsp;line
    npm install npm -g
```

Para verificar se você já tem o NPM instalado e veja a versão instalada, execute o seguinte na linha de comando:

```command&nbsp;line
npm -v
```

Você pode querer usar um Gerenciador de versão do nó para permitir que você alterne entre várias versões do node. js e do NPM, mas isso não é estritamente necessário. Para obter detalhes sobre como fazer isso, [consulte as instruções do NPM](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm).

## <a name="get-office-365"></a>Obter o Office 365

Caso ainda não tenha uma conta do Office 365, obtenha uma assinatura gratuita renovável por 90 dias do Office 365 ingressando no [Programa para Desenvolvedores do Office 365](https://developer.microsoft.com/office/dev-program).

## <a name="install-a-code-editor"></a>Instalar um editor de códigos

Você pode usar qualquer editor de código ou IDE que dê suporte ao desenvolvimento do lado do cliente para criar a web part, como:

- [Visual Studio Code](https://code.visualstudio.com/)
- [Atom](https://atom.io)
- [Webstorm](https://www.jetbrains.com/webstorm)

## <a name="next-steps"></a>Próximas etapas

Tente criar seu próprio suplemento ou use o script Lab para experimentar exemplos internos.

### <a name="create-an-office-add-in"></a>Criar um suplemento do Office

Você pode criar rapidamente um suplemento básico para o Excel, o OneNote, o Outlook, o PowerPoint, o Project ou o Word realizando um [início rápido de 5 minutos](../index.md). Se você já concluiu um início rápido e deseja criar um suplemento um pouco mais complexo, experiente o [tutorial](../index.md).

### <a name="explore-the-apis-with-script-lab"></a>Explorar as APIs com o Script Lab

Explore a biblioteca de amostras internas no [Script Lab](explore-with-script-lab.md) para ter uma ideia dos recursos das APIs JavaScript para Office.

## <a name="see-also"></a>Confira também

- [Criando Suplementos do Office ](../overview/office-add-ins-fundamentals.md)
- [Principais conceitos dos Suplementos do Office](../overview/core-concepts-office-add-ins.md)
- [Desenvolver Suplementos do Office ](../develop/develop-overview.md)
- [Fazer o design de Suplementos do Office](../design/add-in-design.md)
- [Testar e depurar Suplementos do Office](../testing/test-debug-office-add-ins.md)
- [Publicar Suplementos do Office](../publish/publish.md)