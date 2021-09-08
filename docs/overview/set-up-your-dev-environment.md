---
title: Defina seu ambiente de desenvolvimento
description: Configurar seu ambiente de desenvolvedor para criar Office Desempresos.
ms.date: 07/08/2021
localization_priority: Normal
ms.openlocfilehash: efc89b728117e2888cdebd2c5a132047fe662915
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58938760"
---
# <a name="set-up-your-development-environment"></a>Defina seu ambiente de desenvolvimento

Este guia ajuda você a configurar ferramentas para que você possa criar Office de complementos seguindo nossas iniciações rápidas ou tutoriais. Você precisará instalar as ferramentas na lista abaixo. Se você já tiver esses instalados, você estará pronto para iniciar um início rápido, como este Excel React [início rápido.](../quickstarts/excel-quickstart-react.md)

- Node.js
- npm
- Uma Microsoft 365 que inclui a versão de assinatura do Office
- Um editor de código de sua escolha

Este guia supõe que você saiba como usar uma ferramenta de linha de comando.

## <a name="install-nodejs"></a>Instale o Node.js.

Node.js é um tempo de execução JavaScript que você precisará desenvolver Office Descritos modernos.

Instale Node.js baixando a versão recomendada [mais recente do site.](https://nodejs.org) Siga as instruções de instalação do sistema operacional.

## <a name="install-npm"></a>Instalar npm

npm é um registro de software de código aberto do qual baixar os pacotes usados no desenvolvimento Office Desem.

Para instalar npm, execute o seguinte na linha de comando.

```command&nbsp;line
    npm install npm -g
```

Para verificar se você já instalou npm e ver a versão instalada, execute o seguinte na linha de comando.

```command&nbsp;line
npm -v
```

Talvez você queira usar um gerenciador de versão do Nó para permitir que você alterne entre várias versões do Node.js e npm, mas isso não é estritamente necessário. Para obter detalhes sobre como fazer isso, [consulte as instruções do npm.](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm)

## <a name="get-microsoft-365"></a>Obter Microsoft 365

Se você ainda não tiver uma conta de Microsoft 365, poderá obter uma assinatura de 90 dias renováveis de 9 Microsoft 365 0 dias que inclui todos os aplicativos Office, in juntando-se ao programa de desenvolvedor [do Microsoft 365](https://developer.microsoft.com/office/dev-program).

## <a name="install-a-code-editor"></a>Instalar um editor de códigos

Você pode usar qualquer editor de código ou IDE que dê suporte ao desenvolvimento do lado do cliente para criar a web part, como:

- [Visual Studio Code](https://code.visualstudio.com/)
- [Atom](https://atom.io)
- [Webstorm](https://www.jetbrains.com/webstorm)

## <a name="next-steps"></a>Próximas etapas

Tente criar seu próprio complemento ou use Script Lab para experimentar exemplos integrados.

### <a name="create-an-office-add-in"></a>Criar um Suplemento do Office

Você pode criar rapidamente um suplemento básico para o Excel, o OneNote, o Outlook, o PowerPoint, o Project ou o Word realizando um [início rápido de 5 minutos](../index.yml). Se você já concluiu um início rápido e deseja criar um suplemento um pouco mais complexo, experiente o [tutorial](../index.yml).

### <a name="explore-the-apis-with-script-lab"></a>Explorar as APIs com o Script Lab

Explore a biblioteca de amostras internas no [Script Lab](explore-with-script-lab.md) para ter uma ideia dos recursos das APIs JavaScript para Office.

## <a name="see-also"></a>Confira também

- [Principais conceitos dos Suplementos do Office](../overview/core-concepts-office-add-ins.md)
- [Desenvolvimento de Suplementos do Office ](../develop/develop-overview.md)
- [Fazer o design de Suplementos do Office](../design/add-in-design.md)
- [Testar e depurar Suplementos do Office](../testing/test-debug-office-add-ins.md)
- [Publish Office Add-ins](../publish/publish.md)
- [Saiba mais sobre o Programa para Desenvolvedores do Microsoft 365](https://developer.microsoft.com/microsoft-365/dev-program)