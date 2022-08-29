---
title: Criar projetos de suplemento do Office usando o Gerador Yeoman
description: Saiba como criar projetos de suplemento do Office usando o gerador Yeoman para suplementos do Office.
ms.date: 08/19/2022
ms.localizationpriority: high
ms.openlocfilehash: f109c4dbc386a4cc23f72d0c67f9e4904360bba4
ms.sourcegitcommit: 0be4cd0680d638cf96c12263a71af59ff9f51f5a
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/24/2022
ms.locfileid: "67422769"
---
# <a name="create-office-add-in-projects-using-the-yeoman-generator"></a>Criar projetos de suplemento do Office usando o Gerador Yeoman

O [Gerador Yeoman para Suplementos do Office](https://github.com/OfficeDev/generator-office) (também chamado de "Yo Office") é uma ferramenta interativa de linha de comando baseada em Node.js que cria projetos de desenvolvimento de suplementos do Office. Recomendamos que você use essa ferramenta para criar projetos de suplemento, exceto quando quiser que o código do lado do servidor do suplemento esteja em um . Linguagem baseada em NET (como C# ou VB.Net) ou você deseja que o suplemento seja hospedado no IIS (Servidor de Informações da Internet). Em qualquer uma das duas últimas situações, [use o Visual Studio para criar o suplemento](develop-add-ins-visual-studio.md).

Os projetos que a ferramenta cria têm as seguintes características.

- Eles têm uma [configuração npm](https://www.npmjs.com/) padrão que inclui um **arquivo package.json** .
- Eles incluem vários scripts úteis para compilar o projeto, iniciar o servidor, realizar sideload do suplemento no Office e outras tarefas.
- Eles usam [o webpack](https://webpack.js.org/) como um empacotador e um executor de tarefas básico.
- No modo de desenvolvimento, eles são hospedados no localhost pelo webpack-dev-server baseado em Node.js webpack, uma versão orientada para desenvolvimento do [servidor express](http://expressjs.com/) que dá suporte ao recarregamento frequente e recompilamento em alterações.
- Por padrão, todas as dependências são instaladas pela ferramenta, mas você pode adiar a instalação com um argumento de linha de comando.
- Eles incluem um manifesto de suplemento completo.
- Eles têm um suplemento no nível "Olá, Mundo" que está pronto para ser executado assim que a ferramenta for concluída.
- Eles incluem um polyfill e um transcompilador configurado para transpile TypeScript e versões recentes do JavaScript para JavaScript ES5. Esses recursos garantem que o suplemento tenha suporte em todos os runtimes em que os Suplementos do Office podem ser executados, incluindo o Internet Explorer.

> [!TIP]
> Se você quiser se desviar significativamente dessas opções, como usar um executor de tarefas diferente ou um servidor diferente, recomendamos que, ao executar a ferramenta, você escolha a opção [somente manifesto](#manifest-only-option).

## <a name="install-the-generator"></a>Instalar o gerador

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="use-the-tool"></a>Usar a ferramenta

Inicie a ferramenta com o comando a seguir em um prompt do sistema (não em uma janela bash).

```command&nbsp;line
yo office 
```

Muito precisa ser carregado, portanto, pode levar 20 segundos antes que a ferramenta seja iniciada. A ferramenta faz uma série de perguntas. Para alguns, basta digitar uma resposta para o prompt. Para outras pessoas, você recebe uma lista de possíveis respostas. Se receber uma lista, selecione uma e, em seguida, selecione Enter.

A primeira pergunta solicita que você escolha entre seis tipos de projetos. As opções são:

- **Projeto do Painel de Tarefas do Suplemento do Office**
- **Projeto do Painel de Tarefas do Suplemento do Office usando Angular estrutura**
- **Projeto do Painel de Tarefas do Suplemento do Office usando React estrutura**
- **Projeto do Painel de Tarefas do Suplemento do Office que dá suporte ao logon único**
- **Projeto de suplemento do Office que contém apenas o manifesto**
- **Projeto de Suplemento de Funções Personalizadas do Excel**

![Captura de tela mostrando o prompt de tipo de projeto e as possíveis respostas no gerador Yeoman.](../images/yo-office-project-type-prompt.png)

> [!NOTE]
> O **projeto** do Painel de Tarefas do Suplemento do Office que dá suporte à opção de logon único produz um projeto que pode ser usado para ver como o SSO (logon único) funciona em um suplemento. O projeto não pode ser usado como base de um suplemento de produção. Para obter um projeto habilitado para SSO que possa ser uma base de um suplemento de produção, consulte a versão "Completa" de um dos [exemplos de SSO em nosso repositório de exemplos](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth).
>
> O **projeto de suplemento do Office** que contém a opção somente manifesto produz um projeto que contém um manifesto de suplemento básico e um scaffolding mínimo. Para obter mais informações sobre a opção, consulte [a opção Somente manifesto](#manifest-only-option).

A próxima pergunta solicita que você escolha entre **TypeScript** e **JavaScript**. (Essa pergunta será ignorada se você escolher a opção somente manifesto na pergunta anterior.)

![Captura de tela mostrando que o usuário escolheu "Projeto do Painel de Tarefas do Suplemento do Office" para a pergunta anterior e mostra o prompt de idioma e as possíveis respostas, TypeScript e JavaScript, no gerador Yeoman.](../images/yo-office-language-prompt.png)

Em seguida, você será solicitado a dar um nome ao suplemento. O nome especificado será usado no manifesto do suplemento, mas você poderá alterá-lo mais tarde.

![Captura de tela mostrando que o usuário escolheu TypeScript para a pergunta anterior e mostra o prompt do nome do suplemento no gerador Yeoman.](../images/yo-office-name-prompt.png)

Em seguida, você será solicitado a escolher em qual aplicativo do Office o suplemento deve ser executado. Há seis aplicativos possíveis para escolher: **Excel**, **OneNote**, **Outlook**, **PowerPoint**, **Project** e **Word**. Você deve escolher apenas um, mas poderá alterar o manifesto posteriormente para dar suporte aos aplicativos adicionais do Office. A exceção é o Outlook. Um manifesto que dá suporte ao Outlook não pode dar suporte a nenhum outro aplicativo do Office.

![Captura de tela mostrando que o usuário nomeou o projeto "Suplemento Contoso" e mostra o prompt do aplicativo do Office e as possíveis respostas no gerador Yeoman.](../images/yo-office-host-prompt.png)

Depois de responder a essa pergunta, o gerador cria o projeto e instala as dependências. Você pode ver **mensagens WARN** na saída npm na tela. Você pode ignorá-los. Você também pode ver mensagens de que as vulnerabilidades foram encontradas. Você pode ignorá-los por enquanto, mas eventualmente precisará corrigi-los antes que o suplemento seja liberado para produção. Para obter mais informações sobre como corrigir vulnerabilidades, abra o navegador e pesquise "vulnerabilidade npm".

Se a criação for bem-sucedida, você verá um **Parabéns!** na janela comando, seguida por algumas próximas etapas sugeridas. (Se você estiver usando o gerador como parte de um início rápido ou tutorial, ignore as próximas etapas na janela de comando e continue com as instruções no artigo.)

> [!TIP]
> Se você quiser criar o scaffolding de um projeto de Suplemento do Office, mas adiar a instalação das dependências, adicione `--skip-install` a opção ao `yo office` comando. O código a seguir é um exemplo.
>
> ```command&nbsp;line
> yo office --skip-install
> ```
>
> Quando estiver pronto para instalar as dependências, navegue até a pasta raiz do projeto em um prompt de comando e insira `npm install`.

## <a name="manifest-only-option"></a>Opção somente de manifesto

Essa opção cria apenas um manifesto para um suplemento. O projeto resultante não tem um Olá, Mundo suplemento, nenhum dos scripts ou nenhuma das dependências. Use essa opção nos cenários a seguir.

- Você deseja usar ferramentas diferentes das que um projeto gerador Yeoman instala e configura por padrão. Por exemplo, você deseja usar um empacotador, transpilação, executor de tarefas ou servidor de desenvolvimento diferente.
- Você deseja usar uma estrutura de desenvolvimento de aplicativo Web, além Angular ou React, como o Vue.

Para obter um exemplo de como usar o gerador com a opção somente manifesto, consulte [Usar o Vue para criar um suplemento do painel de tarefas do Excel](../quickstarts/excel-quickstart-vue.md).

## <a name="use-command-line-parameters"></a>Usar parâmetros de linha de comando

Você também pode adicionar parâmetros ao `yo office` comando. As duas opções mais comuns são:

- `yo office --details`: isso gerará uma breve ajuda sobre todos os outros parâmetros de linha de comando.
- `yo office --skip-install`: isso impedirá que o gerador instale as dependências.

Para obter referência detalhada sobre os parâmetros de linha de comando, consulte o leiame do gerador no [gerador Yeoman para suplementos do Office](https://github.com/officedev/generator-office).

## <a name="troubleshooting"></a>Solução de problemas

Se você encontrar problemas ao usar a ferramenta, sua primeira etapa deverá ser reinstalá-la para ter certeza de que você tem a versão mais recente. (Consulte [Instalar o gerador](#install-the-generator) para obter detalhes.) Se isso não corrigir o problema, pesquise os problemas do repositório [GitHub](https://github.com/OfficeDev/generator-office/issues) para ver se outra pessoa encontrou o mesmo problema e encontrou uma solução. Se ninguém tiver, [crie um novo problema](https://github.com/OfficeDev/generator-office/issues/new?assignees=&labels=needs+triage&template=bug_report.md&title=).
