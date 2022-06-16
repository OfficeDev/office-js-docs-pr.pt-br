---
title: Tutorial de funções personalizadas do Excel
description: Neste tutorial, você criará um suplemento do Excel que contém uma função personalizada que pode fazer cálculos e solicitar ou transmitir dados da web.
ms.date: 06/10/2022
ms.prod: excel
ms.localizationpriority: high
ms.openlocfilehash: 9550986edcbbed56c69e25e183c304ebe6f6cc07
ms.sourcegitcommit: 4f19f645c6c1e85b16014a342e5058989fe9a3d2
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/15/2022
ms.locfileid: "66091053"
---
# <a name="tutorial-create-custom-functions-in-excel"></a>Tutorial: Criar funções personalizadas no Excel

Funções personalizadas permitem que você adicione novas funções do Excel definindo essas funções em JavaScript como parte de um suplemento. Os usuários do Excel podem acessar funções personalizadas como fariam com qualquer função nativa no Excel, como `SUM()`. Você pode criar funções personalizadas que realizam tarefas simples como cálculos ou tarefas mais complexas, como streaming de dados da web em tempo real em uma planilha.

Neste tutorial, você vai:
> [!div class="checklist"]
> - Crie um suplemento de função personalizada usando o [gerador Yeoman de suplementos do Office](../develop/yeoman-generator-overview.md).
> - Usar uma função personalizada predefinida para realizar um cálculo simples.
> - Criar uma função personalizada que solicita dados da web.
> - Criar uma função personalizada que transmite os dados da web em tempo real.

## <a name="prerequisites"></a>Pré-requisitos

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- Office conectado a uma assinatura Microsoft 365 (incluindo o Office na web).

  > [!NOTE]
  > Se você ainda não tem o Office, poderá [ingressar no programa para desenvolvedores do Microsoft 365](https://developer.microsoft.com/office/dev-program) para obter uma assinatura do Microsoft 365 gratuita e renovável por 90 dias para usar durante o desenvolvimento.

## <a name="create-a-custom-functions-project"></a>Criar um projeto com funções personalizadas

 Para começar, crie do código do projeto para criar o seu suplemento de função personalizada. O [gerador Yeoman para suplementos do Office](../develop/yeoman-generator-overview.md) configura seu projeto com algumas funções personalizadas predefinidas que você pode experimentar. Se você executou a inicialização rápida de funções personalizadas e gerou um projeto, continue usando o projeto e pule para [esta etapa](#create-a-custom-function-that-requests-data-from-the-web).

> [!NOTE]
> Se você recriar o projeto Yo Office, poderá receber um erro porque o cache do Office já tem uma instância de uma função com o mesmo nome. Você pode evitar isso [limpando o cache do Office](../testing/clear-cache.md) antes de executar `npm run start`.

1. [!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

    - **Escolha o tipo de projeto:** `Excel Custom Functions Add-in project`
    - **Escolha o tipo de script:** `JavaScript`
    - **Qual será o nome do suplemento?** `My custom functions add-in`

    :::image type="content" source="../images/yo-office-excel-cf-quickstart.png" alt-text="Captura de tela da interface de linha de comando do gerador do suplemento Yeoman Office para projetos de funções personalizadas.":::

    O gerador Yeoman criará os arquivos do projeto e instalará os componentes Node de suporte.

    [!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

1. Navegue até a pasta raiz do projeto.

    ```command&nbsp;line
    cd "My custom functions add-in"
    ```

1. Compile o projeto.

    ```command&nbsp;line
    npm run build
    ```

    > [!NOTE]
    > Suplementos do Office devem usar HTTPS, não HTTP, mesmo quando você estiver desenvolvendo. Se você for solicitado a instalar um certificado depois de executar `npm run build`, aceite o prompt para instalar o certificado fornecido pelo gerador do Yeoman.

1. Inicie o servidor local da web, que é executado no Node.js. Você pode experimentar o suplemento de função personalizada no Excel.

# <a name="excel-on-windows-or-mac"></a>[Excel para Windows ou Mac](#tab/excel-windows)

Para testar o seu suplemento no Excel para Windows ou Mac, execute o seguinte comando. Quando você executa este comando, o servidor Web local iniciará e o Excel abrirá com o seu suplemento carregado.

```command&nbsp;line
npm run start:desktop
```

[!INCLUDE [alert use https](../includes/alert-use-https.md)]

# <a name="excel-on-the-web"></a>[Excel na Web](#tab/excel-online)

Para testar o suplemento no Excel na Web, execute o seguinte comando. O servidor Web local será iniciado ao executar este comando. Substitua “{url}” pelo URL de um documento do Excel no seu OneDrive ou uma biblioteca do SharePoint para a qual você tenha permissões.

[!INCLUDE [npm start:web command syntax](../includes/start-web-sideload-instructions.md)]

[!INCLUDE [alert use https](../includes/alert-use-https.md)]

---

## <a name="try-out-a-prebuilt-custom-function"></a>Experimente uma função personalizada predefinida

O projeto de funções personalizadas criado contém algumas funções personalizadas predefinidas configuradas no arquivo **src/functions/functions.js**. O arquivo **./manifest.xml** especifica que todas as funções personalizadas pertencem a `CONTOSO` namespace. Você usará o namespace CONTOSO para acessar as funções personalizadas no Excel.

Experimentar a função personalizada `ADD` preenchendo as seguintes etapas no Excel.

1. No Excel, vá para qualquer célula e digite `=CONTOSO`. Observe que o menu de preenchimento automático mostra a lista de todas as funções na `CONTOSO` namespace.

1. Executar a `CONTOSO.ADD` função, com números `10` e `200` como parâmetros de entrada, especificando o valor `=CONTOSO.ADD(10,200)` na célula e pressionando enter.

As `ADD` função personalizada calcula a soma dos dois números que você forneceu e retorna o resultado da **210**.

[!include[Manually register an add-in](../includes/excel-custom-functions-manually-register.md)]

## <a name="create-a-custom-function-that-requests-data-from-the-web"></a>Criar uma função personalizada que solicita dados da web

Integração de dados da Web é uma ótima maneira de ampliar o Excel por meio de funções personalizadas. Em seguida, você criará uma função personalizada chamada `getStarCount` que mostra quantas estrelas um determinado repositório do GitHub tem.

1. No projeto **O meu suplemento de funções personalizadas** localize o arquivo **./src/functions/functions.js** e abra-o no editor de código.

1. Em **function.js**, adicione o código a seguir.

    ```JS
    /**
      * Gets the star count for a given Github repository.
      * @customfunction 
      * @param {string} userName string name of Github user or organization.
      * @param {string} repoName string name of the Github repository.
      * @return {number} number of stars given to a Github repository.
      */
      async function getStarCount(userName, repoName) {
        try {
          //You can change this URL to any web request you want to work with.
          const url = "https://api.github.com/repos/" + userName + "/" + repoName;
          const response = await fetch(url);
          //Expect that status code is in 200-299 range
          if (!response.ok) {
            throw new Error(response.statusText)
          }
            const jsonResponse = await response.json();
            return jsonResponse.watchers_count;
        }
        catch (error) {
          return error;
        }
      }
    ```

1. Execute o seguinte comando para recriar o projeto.

    ```command&nbsp;line
    npm run build
    ```

1. Execute as etapas a seguir (para o Excel na Web, Windows ou Mac) para registrá-lo novamente no Excel. Você deve concluir essas etapas antes que a nova função esteja disponível.

### <a name="excel-on-windows-or-mac"></a>[Excel para Windows ou Mac](#tab/excel-windows)

1. Feche o Excel e abra-o novamente.

1. No Excel, escolha a guia **Inserir** e escolha a seta para baixo localizada à direita de **Meus Suplementos**.

    :::image type="content" source="../images/select-insert.png" alt-text="Captura de tela da faixa de opções Inserir no Excel no Windows, com a seta para baixo Meus suplementos realçada.":::

1. Na lista de suplementos disponíveis, localize a seção **Suplementos do Desenvolvedor** e selecione **Meus suplementos de funções personalizadas** para registrá-lo.

    :::image type="content" source="../images/excel-cf-tutorial-register.png" alt-text="Captura de tela da faixa de opções Inserir no Excel no Windows, com o suplemento funções personalizadas do Excel destacado na lista Meus suplementos.":::

# <a name="excel-on-the-web"></a>[Excel na Web](#tab/excel-online)

1. No Excel, escolha a guia **Inserir** e, em seguida, escolha **Suplementos**.

    :::image type="content" source="../images/excel-cf-online-register-add-in-1.png" alt-text="Captura de tela da faixa de opções Inserir no Excel na web, com o botão Meus suplementos destacado.":::

1. Escolha **Gerenciar Meus suplementos** e selecione **Carregar o Suplemento**.

1. Escolha **Procurar...** e navegue até o diretório raiz do projeto criado pelo gerador Yeoman.

1. Selecione o arquivo **manifest. XML** e escolha **abrir**, escolha **Carregar**.

1. Agora, vamos experimentar a nova função. Na célula **B1**, digite o texto **=CONTOSO. GETSTARCOUNT("OfficeDev", "Excel-Custom-Functions")** e pressione Enter. Você deve ver que o resultado na célula **B1** é o número atual de estrelas fornecido para o [repositório do GitHub de funções personalizadas do Excel](https://github.com/OfficeDev/Excel-Custom-Functions).

---

## <a name="create-a-streaming-asynchronous-custom-function"></a>Criar uma função personalizada assíncrona de streaming

A função `getStarCount` retorna o número de estrelas que um repositório tem em um determinado momento. As funções personalizadas também retornam dados que estão sendo alterados continuamente. Essas funções são chamadas de funções de streaming. Elas devem incluir um parâmetro `invocation` que se refere à célula que chamou a função. O parâmetro `invocation` é usado para atualizar o conteúdo da célula a qualquer momento.  

No exemplo de código a seguir, você perceberá que há duas funções, `currentTime` e `clock`. A função `currentTime` é uma função estática que não usa streaming. Ele retorna a data como uma cadeia de caracteres. A função `clock` usa a função `currentTime` para fornecer o novo horário a cada segundo a uma célula no Excel. Ela usa `invocation.setResult` para fornecer o horário para a célula do Excel e `invocation.onCanceled` para controlar o que acontece quando a função é cancelada. 

O projeto **Meus suplementos de funções personalizadas** já contém as duas funções a seguir no arquivo **./src/functions/functions.js**.

```JS
/**
 * Returns the current time
 * @returns {string} String with the current time formatted for the current locale.
 */
function currentTime() {
  return new Date().toLocaleTimeString();
}
    
/**
 * Displays the current time once a second
 * @customfunction
 * @param {CustomFunctions.StreamingInvocation<string>} invocation Custom function invocation
 */
function clock(invocation) {
  const timer = setInterval(() => {
    const time = currentTime();
    invocation.setResult(time);
  }, 1000);
    
  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}
```

Para experimentar as funções, digite o texto **=CONTOSO. CLOCK()** na célula **C1** e pressione Enter. Você deverá ver a data atual, que transmite uma atualização a cada segundo. Embora esse relógio seja um cronômetro em um loop, você pode usar a mesma ideia para definir um cronômetro em funções mais complexas que fazem solicitações da Web para dados em tempo real.

## <a name="next-steps"></a>Próximas etapas

Parabéns! Neste tutorial, você criou um novo projeto de funções personalizadas, experimentou uma função predefinida, criou uma função personalizada que solicita dados da Web e criou uma função personalizada que transmite dados. Em seguida, você pode modificar seu projeto para usar um tempo de execução compartilhado, facilitando a interação com o painel de tarefas. Siga as etapas no seguinte artigo.

> [!div class="nextstepaction"]
> [Configure seu suplemento para usar um tempo de execução compartilhado](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
