---
title: Defina seu ambiente de desenvolvimento
description: Configurar seu ambiente de desenvolvedor para criar Os Complementos do Office.
ms.date: 02/09/2021
localization_priority: Normal
ms.openlocfilehash: 1dd0cc6bb035a0274e36fe9916dcd2481bdf0b39
ms.sourcegitcommit: fefc279b85e37463413b6b0e84c880d9ed5d7ac3
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/12/2021
ms.locfileid: "50234125"
---
# <a name="set-up-your-development-environment"></a>Defina seu ambiente de desenvolvimento

Este guia ajuda você a configurar ferramentas para que você possa criar Os Complementos do Office seguindo nossos inícios ou tutoriais rápidos. Você precisará instalar as ferramentas na lista abaixo. Se você já tiver instalado, você está pronto para começar um início rápido, como este [excel React início rápido.](../quickstarts/excel-quickstart-react.md)

- Node.js
- npm
- Uma conta do Microsoft 365 que inclui a versão de assinatura do Office
- Um editor de código de sua escolha

Este guia assume que você sabe como usar uma ferramenta de linha de comando. 

## <a name="install-nodejs"></a>Instale o Node.js.

Node.js é um tempo de execução JavaScript que você precisará para desenvolver Complementos modernos do Office.

Instale Node.js [baixando a versão mais recente recomendada do site.](https://nodejs.org) Siga as instruções de instalação do sistema operacional.

## <a name="install-npm"></a>Instalar npm

O npm é um registro de software aberto do qual baixar os pacotes usados no desenvolvimento de Complementos do Office.

Para instalar o npm, execute o seguinte na linha de comando:

```command&nbsp;line
    npm install npm -g
```

Para verificar se você já tem o npm instalado e ver a versão instalada, execute o seguinte na linha de comando:

```command&nbsp;line
npm -v
```

Talvez você queira usar um gerenciador de versão do Node para permitir que você alternar entre várias versões do Node.js e npm, mas isso não é estritamente necessário. Para obter detalhes sobre como fazer isso, [consulte as instruções do npm.](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm)

## <a name="get-microsoft-365"></a>Obter o Microsoft 365

Se você ainda não tiver uma conta do Microsoft 365, poderá obter uma assinatura gratuita e renovável de 90 dias do Microsoft 365 que inclua todos os aplicativos do Office in joining ao programa de desenvolvedores [do Microsoft 365.](https://developer.microsoft.com/office/dev-program)

## <a name="install-a-code-editor"></a>Instalar um editor de códigos

Você pode usar qualquer editor de código ou IDE que dê suporte ao desenvolvimento do lado do cliente para criar a web part, como:

- [Visual Studio Code](https://code.visualstudio.com/)
- [Atom](https://atom.io)
- [Webstorm](https://www.jetbrains.com/webstorm)

## <a name="next-steps"></a>Próximas etapas

Tente criar seu próprio add-in ou usar o Script Lab para experimentar exemplos integrados.

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