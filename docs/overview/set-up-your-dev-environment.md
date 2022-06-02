---
title: Defina seu ambiente de desenvolvimento
description: Configure seu ambiente de desenvolvedor para criar Office suplementos.
ms.date: 05/27/2022
ms.localizationpriority: medium
ms.openlocfilehash: 01b9fe0aff2696a521266bb3175ea0f61d891aa4
ms.sourcegitcommit: 35e7646c5ad0d728b1b158c24654423d999e0775
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/02/2022
ms.locfileid: "65833882"
---
# <a name="set-up-your-development-environment"></a>Defina seu ambiente de desenvolvimento

Este guia ajuda você a configurar ferramentas para que você possa criar Office suplementos seguindo nossos guias de início rápido ou tutoriais. Se você já tiver esses recursos instalados, estará pronto para começar rapidamente, como este Excel React [início rápido](../quickstarts/excel-quickstart-react.md).

## <a name="get-microsoft-365"></a>Obter Microsoft 365

Você precisa de uma Microsoft 365 conta. Você pode obter uma assinatura renovável gratuita de 90 dias Microsoft 365 que inclui todos os aplicativos Office ingressando no programa [Microsoft 365 desenvolvedor.](https://developer.microsoft.com/office/dev-program)

## <a name="install-the-environment"></a>Instalar o ambiente

Há dois tipos de ambientes de desenvolvimento para escolher. O scaffolding de projetos de suplemento Office criados nos dois ambientes é diferente, portanto, se várias pessoas estiverem trabalhando em um projeto de suplemento, todas elas deverão usar o mesmo ambiente. 

- **Node.js ambiente**: recomendado. Nesse ambiente, suas ferramentas são instaladas e executadas em uma linha de comando. O lado do servidor da parte do aplicativo Web do suplemento é escrito em JavaScript ou TypeScript e é hospedado em um Node.js runtime. Há muitas ferramentas úteis de desenvolvimento de suplementos nesse ambiente, como um linter Office e um empacotador/executor de tarefas chamado WebPack. A ferramenta de criação e scaffolding do projeto, Yo Office, é atualizada com frequência.
- **Visual Studio** ambiente: escolha esse ambiente somente se o computador de desenvolvimento for Windows e você quiser desenvolver o lado do servidor do suplemento com uma linguagem e estrutura baseadas em .NET, como ASP.NET. Os modelos de projeto de suplemento Visual Studio são atualizados com a mesma frequência que os modelos de projeto Node.js ambiente. O código do lado do cliente não pode ser depurado com o depurador Visual Studio interno, mas você pode depurar o código do lado do cliente com as ferramentas de desenvolvimento do navegador. Mais informações posteriormente na **guia Visual Studio ambiente**.

> [!NOTE]
> Visual Studio para Mac não inclui os modelos de scaffolding de projeto para suplementos do Office, portanto, se o computador de desenvolvimento for um Mac, você deverá trabalhar com o ambiente de Node.js.

Selecione a guia para o ambiente escolhido. 

# <a name="nodejs-environment"></a>[Node.js ambiente](#tab/yeomangenerator)

As principais ferramentas a serem instaladas são:

- Node.js
- npm
- Um editor de código de sua escolha
- Yo Office
- O linter Office JavaScript

Este guia pressupõe que você saiba como usar uma ferramenta de linha de comando.

### <a name="install-nodejs-and-npm"></a>Instalar Node.js e npm

Node.js é um runtime do JavaScript que você usa para desenvolver suplementos Office modernos.

Instale Node.js [baixando a versão recomendada mais recente do site](https://nodejs.org). Siga as instruções de instalação do sistema operacional.

npm é um código aberto de software do qual baixar os pacotes usados no desenvolvimento Office Suplementos. Normalmente, ele é instalado automaticamente quando você instala Node.js. Para verificar se você já npm instalado e ver a versão instalada, execute o seguinte na linha de comando.

```command&nbsp;line
npm -v
```

Se, por qualquer motivo, você quiser instalá-lo manualmente, execute o seguinte na linha de comando.

```command&nbsp;line
npm install npm -g
```

> [!TIP]
> Talvez você queira usar um gerenciador de versões do Node para permitir que você alterne entre várias versões do Node.js e npm, mas isso não é estritamente necessário. Para obter detalhes sobre como fazer isso, [consulte npm instruções do usuário](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm).

### <a name="install-a-code-editor"></a>Instalar um editor de códigos

Você pode usar qualquer editor de código ou IDE que dê suporte ao desenvolvimento do lado do cliente para criar a web part, como:

- [Visual Studio Code](https://code.visualstudio.com/) (recomendado)
- [Atom](https://atom.io)
- [Webstorm](https://www.jetbrains.com/webstorm)

### <a name="install-the-yeoman-generator-mdash-yo-office"></a>Instalar o gerador Yeoman &mdash; Yo Office

A ferramenta de criação e scaffolding do projeto é [o gerador Yeoman para Office suplementos](../develop/yeoman-generator-overview.md), muito conhecidos como **Yo Office**. Você precisa instalar a versão mais recente do [Yeoman](https://github.com/yeoman/yo) e do Yo Office. Para instalar essas ferramentas globalmente, execute o comando a seguir por meio do prompt de comando.

  ```command&nbsp;line
  npm install -g yo generator-office
  ```

### <a name="install-and-use-the-office-javascript-linter"></a>Instalar e usar o linter Office JavaScript

A Microsoft fornece um linter JavaScript para ajudá-lo a detectar erros comuns ao usar a biblioteca Office JavaScript. Para instalar o linter, execute os dois comandos a seguir (depois de instalar Node.js [e npm](#install-nodejs-and-npm)).

```command&nbsp;line
npm install office-addin-lint --save-dev
npm install eslint-plugin-office-addins --save-dev
```

Se você criar um projeto Office suplemento com o gerador [Yeoman para Office suplementos](../develop/yeoman-generator-overview.md), o restante da configuração será feito para você. Execute o linter com o comando a seguir no terminal de um editor, como Visual Studio Code, ou em um prompt de comando. Os problemas encontrados pelo linter aparecem no terminal ou no prompt e também aparecem diretamente no código quando você está usando um editor que dá suporte a mensagens de linter, como Visual Studio Code. (Para obter informações sobre como instalar o gerador Yeoman, consulte o gerador [Yeoman para Office suplementos](../develop/yeoman-generator-overview.md).)

```command&nbsp;line
npm run lint
```

Se o projeto de suplemento tiver sido criado de outra maneira, execute as etapas a seguir.

1. Na raiz do projeto, crie um arquivo de texto chamado **.eslintrc.json**, se ainda não houver um. Verifique se ele tem propriedades nomeadas `plugins` e `extends`, ambos do tipo matriz. A `plugins` matriz deve incluir `"office-addins"` e a `extends` matriz deve incluir `"plugin:office-addins/recommended"`. Apresentamos um exemplo simples a seguir. O **arquivo .eslintrc.json** pode ter propriedades adicionais e membros adicionais das duas matrizes.

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

1. Na raiz do projeto, abra o **arquivo package.json** `scripts` e verifique se a matriz tem o membro a seguir.

   ```json
   "lint": "office-addin-lint check",
   ```

1. Execute o linter com o comando a seguir no terminal de um editor, como Visual Studio Code, ou em um prompt de comando. Os problemas encontrados pelo linter aparecem no terminal ou no prompt e também aparecem diretamente no código quando você está usando um editor que dá suporte a mensagens de linter, como Visual Studio Code.

   ```command&nbsp;line
   npm run lint
   ```

# <a name="visual-studio-environment"></a>[Visual Studio ambiente](#tab/visualstudio)

### <a name="install-visual-studio"></a>Instalar o Visual Studio

Se você não tiver o Visual Studio 2017 (para Windows) ou posterior instalado, instale a versão mais recente do [Visual Studio Downloads](https://visualstudio.microsoft.com/downloads/). Inclua a carga de trabalho **de Office/SharePoint** de desenvolvimento quando o instalador solicitar que você especifique cargas de trabalho. Outras cargas de trabalho que podem ser necessárias são ferramentas de desenvolvimento da **Web para .NET**, **JavaScript e suporte à linguagem TypeScript** (para codificar o lado do cliente do suplemento) e cargas de trabalho ASP.NET relacionadas.

> [!TIP]
> A partir do verão de 2022, os esquemas XML para o manifesto do suplemento Office instalados com Visual Studio não são a versão mais recente. Isso pode afetar os suplementos, dependendo de quais recursos de suplemento eles usam. Portanto, talvez seja necessário atualizar os esquemas XML para o manifesto. Para obter mais informações, consulte [Erros de validação de esquema de manifesto Visual Studio projetos](../testing/troubleshoot-development-errors.md#manifest-schema-validation-errors-in-visual-studio-projects).

> [!NOTE]
> Para obter informações sobre como depurar o código do lado do cliente quando você estiver usando o ambiente Visual Studio, consulte [Depurar Office Suplementos no Visual Studio](../develop/debug-office-add-ins-in-visual-studio.md). Depure o código do lado do servidor da mesma maneira que qualquer aplicativo Web criado Visual Studio. Consulte [o lado do cliente ou do servidor](../testing/debug-add-ins-overview.md#server-side-or-client-side).

---

## <a name="install-script-lab"></a>Instalar Script Lab

Script Lab é uma ferramenta para criar protótipos rápidos de código que chama as APIs Office Biblioteca JavaScript. Script Lab é um Office suplemento e pode ser instalado do AppSource [em Script Lab](https://appsource.microsoft.com/marketplace/apps?search=script%20lab&page=1). Há uma versão para Excel, PowerPoint Word e uma versão separada para Outlook. Para obter informações sobre como usar Script Lab, consulte [Explorar Office API JavaScript usando Script Lab](explore-with-script-lab.md).

## <a name="next-steps"></a>Próximas etapas

Tente criar seu próprio suplemento ou use [Script Lab](explore-with-script-lab.md) para experimentar exemplos internos.

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