---
title: Defina seu ambiente de desenvolvimento
description: Configurar seu ambiente de desenvolvedor para criar Office Desempresos.
ms.date: 10/26/2021
ms.localizationpriority: medium
ms.openlocfilehash: ad1fc265640b6fb5931ba2086cc61784e94365c1
ms.sourcegitcommit: 4a7b9b9b359d51688752851bf3b41b36f95eea00
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/22/2022
ms.locfileid: "63710920"
---
# <a name="set-up-your-development-environment"></a>Defina seu ambiente de desenvolvimento

Este guia ajuda você a configurar ferramentas para que você possa criar Office Desindados seguindo nossas iniciações rápidas ou tutoriais. Você precisará instalar as ferramentas na lista abaixo. Se você já tiver esses instalados, você estará pronto para iniciar um início rápido, como este [Excel React início rápido](../quickstarts/excel-quickstart-react.md).

- Node.js
- npm
- Uma Microsoft 365 que inclui a versão de assinatura do Office
- Um editor de código de sua escolha
- O Office JavaScript

Este guia supõe que você saiba como usar uma ferramenta de linha de comando.

## <a name="install-nodejs"></a>Instale o Node.js.

Node.js é um tempo de execução JavaScript que você precisará desenvolver Office Descritos modernos.

Instale Node.js [baixando a versão recomendada mais recente em seu site](https://nodejs.org). Siga as instruções de instalação do sistema operacional.

## <a name="install-npm"></a>Instalar npm

npm é um registro de software de código aberto a partir do qual baixar os pacotes usados no desenvolvimento Office de complementos.

Para instalar npm, execute o seguinte na linha de comando.

```command&nbsp;line
    npm install npm -g
```

Para verificar se você já instalou npm e ver a versão instalada, execute o seguinte na linha de comando.

```command&nbsp;line
npm -v
```

Talvez você queira usar um gerenciador de versão do Nó para permitir que você alterne entre várias versões do Node.js e npm, mas isso não é estritamente necessário. Para obter detalhes sobre como fazer isso, [consulte as instruções do npm](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm).

## <a name="get-microsoft-365"></a>Obter Microsoft 365

Se você ainda não tiver uma conta de Microsoft 365, poderá obter uma assinatura de Microsoft 365 renovável de 90 dias, que inclui todos os aplicativos Office, in juntando-se ao programa de desenvolvedor [Microsoft 365](https://developer.microsoft.com/office/dev-program) de 90 dias.

## <a name="install-a-code-editor"></a>Instalar um editor de códigos

Você pode usar qualquer editor de código ou IDE que dê suporte ao desenvolvimento do lado do cliente para criar a web part, como:

- [Visual Studio Code](https://code.visualstudio.com/)
- [Atom](https://atom.io)
- [Webstorm](https://www.jetbrains.com/webstorm)

## <a name="install-and-use-the-office-javascript-linter"></a>Instalar e usar o Office JavaScript

A Microsoft fornece um linter JavaScript para ajudá-lo a capturar erros comuns ao usar a biblioteca Office JavaScript. Para instalar o linter, execute os dois comandos a seguir (depois de instalar [Node.jse ](#install-nodejs) [npm](#install-npm)).

```command&nbsp;line
npm install office-addin-lint --save-dev
npm install eslint-plugin-office-addins --save-dev
```

Se você criar um projeto de Office com o gerador [Yeoman](../develop/yeoman-generator-overview.md) para Office de complementos, o restante da instalação será feito para você. Execute o linter com o seguinte comando no terminal de um editor, como Visual Studio Code, ou em um prompt de comando. Os problemas encontrados pelo linter aparecem no terminal ou no prompt e também aparecem diretamente no código quando você está usando um editor que dá suporte a mensagens de linter, como Visual Studio Code. (Para obter informações sobre como instalar o gerador Yeoman, consulte [Yeoman generator for Office Add-ins](../develop/yeoman-generator-overview.md).)

```command&nbsp;line
npm run lint
```

Se o projeto do seu add-in foi criado de outra maneira, tome as etapas a seguir.

1. Na raiz do projeto, crie um arquivo de texto chamado **.eslintrc.json**, se ainda não houver um. Certifique-se de que ele tenha propriedades nomeadas `plugins` e `extends`, ambas de matriz de tipo. A `plugins` matriz deve incluir e `"office-addins"` a `extends` matriz deve incluir `"plugin:office-addins/recommended"`. Apresentamos um exemplo simples a seguir. Seu **arquivo .eslintrc.json** pode ter propriedades adicionais e membros adicionais das duas matrizes.

   ```json
   {
     "plugins": [
       "office-addins"
     ],
     "extends": [
       "plugin:office-addins/recommended"
     ]
   }
   ```

1. Na raiz do projeto, abra o **arquivo package.json** `scripts` e certifique-se de que a matriz tenha o membro a seguir.

   ```json
   "lint": "office-addin-lint check",
   ```

1. Execute o linter com o seguinte comando no terminal de um editor, como Visual Studio Code, ou em um prompt de comando. Os problemas encontrados pelo linter aparecem no terminal ou no prompt e também aparecem diretamente no código quando você está usando um editor que dá suporte a mensagens de linter, como Visual Studio Code.

   ```command&nbsp;line
   npm run lint
   ```

## <a name="next-steps"></a>Próximas etapas

Tente criar seu próprio add-in ou use Script Lab para experimentar exemplos integrados.

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