---
title: Defina seu ambiente de desenvolvimento
description: Configure seu ambiente de desenvolvedor para criar suplementos do Office.
ms.date: 09/01/2022
ms.localizationpriority: medium
ms.openlocfilehash: 4e03ea7f55786107354f9d5a92e0cb30ffb559ec
ms.sourcegitcommit: 889d23061a9413deebf9092d675655f13704c727
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/07/2022
ms.locfileid: "67615997"
---
# <a name="set-up-your-development-environment"></a>Defina seu ambiente de desenvolvimento

Este guia ajuda você a configurar ferramentas para que você possa criar Suplementos do Office seguindo nossos guias de início rápido ou tutoriais. Se você já tiver essas configurações instaladas, estará pronta para começar um início rápido, como este [Excel React início rápido](../quickstarts/excel-quickstart-react.md).

## <a name="get-microsoft-365"></a>Obter o Microsoft 365

Você precisa de uma conta do Microsoft 365. Você pode obter uma assinatura gratuita e renovável do Microsoft 365 de 90 dias que inclui todos os aplicativos do Office ingressando no programa para desenvolvedores do [Microsoft 365](https://developer.microsoft.com/office/dev-program).

## <a name="install-the-environment"></a>Instalar o ambiente

Há dois tipos de ambientes de desenvolvimento para escolher. O scaffolding de projetos de Suplementos do Office criados nos dois ambientes é diferente, portanto, se várias pessoas estiverem trabalhando em um projeto de suplemento, todas elas deverão usar o mesmo ambiente. 

- **Node.js ambiente**: recomendado. Nesse ambiente, suas ferramentas são instaladas e executadas em uma linha de comando. O lado do servidor da parte do aplicativo Web do suplemento é escrito em JavaScript ou TypeScript e é hospedado em um Node.js runtime. Há muitas ferramentas úteis de desenvolvimento de suplementos nesse ambiente, como um linter do Office e um empacotador/executor de tarefas chamado WebPack. A ferramenta de criação e scaffolding do projeto, Yo Office, é atualizada com frequência.
- **Ambiente do Visual Studio**: escolha esse ambiente somente se o computador de desenvolvimento for Windows e você quiser desenvolver o lado do servidor do suplemento com uma estrutura e linguagem baseada em .NET, como ASP.NET. Os modelos de projeto de suplemento no Visual Studio não são atualizados com a mesma frequência que os modelos de projeto Node.js ambiente. O código do lado do cliente não pode ser depurado com o depurador interno do Visual Studio, mas você pode depurar o código do lado do cliente com as ferramentas de desenvolvimento do navegador. Mais informações posteriormente na guia **de ambiente do Visual Studio** .

> [!NOTE]
> Visual Studio para Mac não inclui os modelos de scaffolding de projeto para Suplementos do Office, portanto, se o computador de desenvolvimento for um Mac, você deverá trabalhar com o ambiente Node.js ambiente.

Selecione a guia para o ambiente escolhido. 

# <a name="nodejs-environment"></a>[Node.js ambiente](#tab/yeomangenerator)

As principais ferramentas a serem instaladas são:

- Node.js
- npm
- Um editor de código de sua escolha
- Yo Office
- O linter JavaScript do Office

Este guia pressupõe que você saiba como usar uma ferramenta de linha de comando.

### <a name="install-nodejs-and-npm"></a>Instalar Node.js e npm

Node.js é um runtime do JavaScript que você usa para desenvolver suplementos modernos do Office.

Instale Node.js [baixando a versão recomendada mais recente do site](https://nodejs.org). Siga as instruções de instalação do sistema operacional.

npm é um código aberto de software do qual baixar os pacotes usados no desenvolvimento de Suplementos do Office. Normalmente, ele é instalado automaticamente quando você instala Node.js. Para verificar se você já tem o npm instalado e ver a versão instalada, execute o seguinte na linha de comando.

```command&nbsp;line
npm -v
```

Se, por qualquer motivo, você quiser instalá-lo manualmente, execute o seguinte na linha de comando.

```command&nbsp;line
npm install npm -g
```

> [!TIP]
> Talvez você queira usar um gerenciador de versões do Node para permitir que você alterne entre várias versões do Node.js e npm, mas isso não é estritamente necessário. Para obter detalhes sobre como fazer isso, [consulte as instruções do npm](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm).

### <a name="install-a-code-editor"></a>Instalar um editor de códigos

Você pode usar qualquer editor de código ou IDE que dê suporte ao desenvolvimento do lado do cliente para criar a web part, como:

- [Visual Studio Code](https://code.visualstudio.com/) (recomendado)
- [Atom](https://atom.io)
- [Webstorm](https://www.jetbrains.com/webstorm)

### <a name="install-the-yeoman-generator-mdash-yo-office"></a>Instalar o gerador Yeoman Yo &mdash; Office

A ferramenta de criação e scaffolding do projeto é o [gerador Yeoman para Suplementos do Office](../develop/yeoman-generator-overview.md), conhecido afetuosamente como **Yo Office**. Você precisa instalar a versão mais recente do [Yeoman](https://github.com/yeoman/yo) e do Yo Office. Para instalar essas ferramentas globalmente, execute o seguinte comando por meio do prompt de comando.

  ```command&nbsp;line
  npm install -g yo generator-office
  ```

### <a name="install-and-use-the-office-javascript-linter"></a>Instalar e usar o linter JavaScript do Office

A Microsoft fornece um linter JavaScript para ajudá-lo a detectar erros comuns ao usar a biblioteca JavaScript do Office. Para instalar o linter, execute os dois comandos a seguir (depois de instalar Node.js [e npm](#install-nodejs-and-npm)).

```command&nbsp;line
npm install office-addin-lint --save-dev
npm install eslint-plugin-office-addins --save-dev
```

Se você criar um projeto de Suplemento do Office com o gerador [Yeoman](../develop/yeoman-generator-overview.md) para a ferramenta suplementos do Office, o restante da configuração será feito para você. Execute o linter com o comando a seguir no terminal de um editor, como Visual Studio Code, ou em um prompt de comando. Os problemas encontrados pelo linter aparecem no terminal ou no prompt e também aparecem diretamente no código quando você está usando um editor que dá suporte a mensagens de linter, como Visual Studio Code. (Para obter informações sobre como instalar o gerador Yeoman, consulte [o gerador Yeoman para suplementos do Office](../develop/yeoman-generator-overview.md).)

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

# <a name="visual-studio-environment"></a>[Ambiente do Visual Studio](#tab/visualstudio)

### <a name="install-visual-studio"></a>Instalar o Visual Studio

Se você não tiver o Visual Studio 2017 (para Windows) ou posterior instalado, instale a versão mais recente dos [Downloads do Visual Studio](https://visualstudio.microsoft.com/downloads/). Inclua a carga de trabalho de **desenvolvimento do Office/SharePoint** quando o instalador solicitar que você especifique cargas de trabalho. Outras cargas de trabalho que podem ser necessárias são ferramentas de desenvolvimento da **Web para .NET**, **JavaScript e suporte à linguagem TypeScript** (para codificar o lado do cliente do suplemento) e cargas de trabalho ASP.NET relacionadas.

> [!TIP]
> A partir de junho de 2022, os esquemas XML para o manifesto do Suplemento do Office instalados com o Visual Studio não são a versão mais recente. Isso pode afetar os suplementos, dependendo de quais recursos de suplemento eles usam. Portanto, talvez seja necessário atualizar os esquemas XML para o manifesto. Para obter mais informações, consulte [Erros de validação de esquema de manifesto em projetos do Visual Studio](../testing/troubleshoot-development-errors.md#manifest-schema-validation-errors-in-visual-studio-projects).

> [!NOTE]
> Para obter informações sobre como depurar código do lado do cliente quando você estiver usando o ambiente do Visual Studio, consulte [Depurar suplementos do Office no Visual Studio](../develop/debug-office-add-ins-in-visual-studio.md). Depure o código do lado do servidor da mesma maneira que qualquer aplicativo Web criado no Visual Studio. Consulte [o lado do cliente ou do servidor](../testing/debug-add-ins-overview.md#server-side-or-client-side).

---

## <a name="install-script-lab"></a>Instalar Script Lab

Script Lab é uma ferramenta para criar rapidamente o código que chama as APIs da Biblioteca JavaScript do Office. Script Lab é um suplemento do Office e pode ser instalado do AppSource em [Script Lab](https://appsource.microsoft.com/marketplace/apps?search=script%20lab&page=1). Há uma versão para Excel, PowerPoint e Word e uma versão separada para o Outlook. Para obter informações sobre como usar Script Lab, consulte [Explorar a API JavaScript do Office usando Script Lab](explore-with-script-lab.md).

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