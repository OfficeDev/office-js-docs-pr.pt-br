---
title: Tutorial de funções personalizadas do Excel
description: Neste tutorial, você criará um suplemento do Excel que contém uma função personalizada que pode executar cálculos e solicitar ou transmitir dados da web.
ms.date: 06/27/2019
ms.prod: excel
ms.topic: tutorial
localization_priority: Normal
ms.openlocfilehash: 1aa05581d1b0dfb1f5affa019e51b84126c8d199
ms.sourcegitcommit: 90c2d8236c6b30d80ac2b13950028a208ef60973
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/02/2019
ms.locfileid: "35454719"
---
# <a name="tutorial-create-custom-functions-in-excel"></a>Tutorial: Criar funções personalizadas no Excel

Funções personalizadas permitem que você adicione novas funções do Excel definindo essas funções em JavaScript como parte de um suplemento. Os usuários do Excel podem acessar funções personalizadas como fariam com qualquer função nativa no Excel, como `SUM()`. Você pode criar funções personalizadas que realizam tarefas simples como cálculos ou tarefas mais complexas, como streaming de dados da web em tempo real em uma planilha.

Neste tutorial, você vai:
> [!div class="checklist"]
> * Crie um suplemento de função personalizada usando o [gerador Yeoman de suplementos do Office](https://www.npmjs.com/package/generator-office). 
> * Usar uma função personalizada predefinida para realizar um cálculo simples.
> * Criar uma função personalizada que solicita dados da web.
> * Criar uma função personalizada que transmite os dados da web em tempo real.

## <a name="prerequisites"></a>Pré-requisitos

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

* Excel no Windows (versão 1904 ou posterior, conectada à assinatura do Office 365) ou na Web

## <a name="create-a-custom-functions-project"></a>Criar um projeto com funções personalizadas

 Para começar, você criará o projeto de código para criar o suplemento função personalizada. O [gerador Yeoman para suplementos do Office](https://www.npmjs.com/package/generator-office) configurará seu projeto com algumas funções personalizadas predefinidas que você pode experimentar. Se você já tiver executado o início rápido de funções personalizadas e gerado um projeto, continue a usar esse projeto e pule para [esta etapa](#create-a-custom-function-that-requests-data-from-the-web) .

1. Execute o comando a seguir e responda aos prompts da seguinte forma.
    
    ```command&nbsp;line
    yo office
    ```
    
    * **Escolha o tipo de projeto:** `Excel Custom Functions Add-in project`
    * **Escolha o tipo de script:** `JavaScript`
    * **Qual será o nome do suplemento?** `starcount`

    ![O gerador Yeoman para suplementos do Office solicita funções personalizadas](../images/starcountPrompt.png)
    
    O gerador Yeoman criará os arquivos do projeto e instalará os componentes Node de suporte.

2. Navegue até a pasta raiz do projeto.
    
    ```command&nbsp;line
    cd starcount
    ```

3. Compile o projeto.
    
    ```command&nbsp;line
    npm run build
    ```

    > [!NOTE]
    > Os Suplementos do Office devem usar HTTPS, e não HTTP, mesmo durante o desenvolvimento. Se você for solicitado a instalar um certificado após executar `npm run build`, aceite a solicitação para instalar o certificado que o gerador do Yeoman fornecer.

4. Inicie o servidor local da web, que é executado no Node. Você pode experimentar o suplemento função personalizada no Excel na Web ou no Windows.

# <a name="excel-on-windowstabexcel-windows"></a>[Excel no Windows](#tab/excel-windows)

Para testar seu suplemento no Excel no Windows, execute o seguinte comando. Quando você executar este comando, o servidor Web local será iniciado e o Excel será aberto com o seu suplemento carregado.

```command&nbsp;line
npm run start:desktop
```

# <a name="excel-on-the-webtabexcel-online"></a>[Excel na Web](#tab/excel-online)

Para testar seu suplemento no Excel em um navegador, execute o seguinte comando. Quando você executa este comando, o servidor Web local iniciará.

```command&nbsp;line
npm run start:web
```

Para usar seu suplemento de funções personalizadas, abra uma nova pasta de trabalho no Excel na Web. Nesta pasta de trabalho, conclua as seguintes etapas para Sideload seu suplemento.

1. No Excel, escolha a guia **Inserir** e, em seguida, escolha **suplementos**.

   ![Inserir faixa de opções no Excel na Web com o ícone meus suplementos realçado](../images/excel-cf-online-register-add-in-1.png)
   
2. Escolha **Gerenciar Meus suplementos** e selecione **Carregar o Suplemento**.

3. Escolha **Procurar... ** e navegue até o diretório raiz do projeto criado pelo gerador Yeoman.

4. Selecione o arquivo **manifest. XML** e escolha **aberto**, escolha **Carregar**.

--- 
    
## <a name="try-out-a-prebuilt-custom-function"></a>Experimente uma função personalizada predefinida

O projeto de funções personalizadas que você criou contém algumas funções personalizadas predefinidas, definidas no arquivo **./src/Functions/functions.js** . O arquivo **./manifest.xml** especifica que todas as funções personalizadas pertencem a `CONTOSO` namespace. Você usará o namespace CONTOSO para acessar as funções personalizadas no Excel.

Em seguida você vai experimentar a função personalizada `ADD` preenchendo as seguintes etapas:

1. No Excel, vá para qualquer célula e digite `=CONTOSO`. Observe que o menu de preenchimento automático mostra a lista de todas as funções na `CONTOSO` namespace.

2. Executar a `CONTOSO.ADD` função, com números `10` e `200` como parâmetros de entrada, especificando o valor `=CONTOSO.ADD(10,200)` na célula e pressionando enter.

As `ADD` função personalizada calcula a soma dos dois números que você forneceu e retorna o resultado da **210**.

## <a name="create-a-custom-function-that-requests-data-from-the-web"></a>Criar uma função personalizada que solicita dados da web

Integração de dados da Web é uma ótima maneira de ampliar o Excel por meio de funções personalizadas. Em seguida, você criará uma função `getStarCount` personalizada chamada que mostra quantas estrelas um determinado repositório do GitHub possui.

1. No projeto **starcount** , localize o arquivo **./src/Functions/functions.js** e abra-o no editor de código. 

2. Em **Function. js**, adicione o seguinte código: 

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
    CustomFunctions.associate("GETSTARCOUNT", getStarCount);
```

O `CustomFunctions.associate` código associa a `id` da função com o endereço de função da `getStarCount` em JavaScript para que o Excel possa ligar para a função.

3. Execute o seguinte comando para recriar o projeto.

    ```command&nbsp;line
    npm run build
    ```

4. Complete as etapas a seguir (para o Excel na Web ou Windows) para registrar novamente o suplemento no Excel. Você deve concluir estas etapas para que a nova função esteja disponível.

### <a name="excel-on-windowstabexcel-windows"></a>[Excel no Windows](#tab/excel-windows)

1. Feche o Excel e abra novamente o Excel.

2. No Excel, escolha a guia **Inserir** e, em seguida, escolha a seta para baixo localizada à direita de **meus**suplementos.  ![Inserir faixa de opções no Excel no Windows com a seta meus suplementos realçada](../images/select-insert.png)

3. Na lista de suplementos disponíveis, encontre a seção suplementos do **desenvolvedor** e selecione o suplemento do **starcount** para registrá-lo.
    ![Inserir faixa de opções no Excel no Windows com o suplemento funções personalizadas do Excel realçado na lista meus suplementos](../images/list-starcount.png)


# <a name="excel-on-the-webtabexcel-online"></a>[Excel na Web](#tab/excel-online)

1. No Excel, escolha a guia **Inserir** e, em seguida, escolha **suplementos**.  ![Inserir faixa de opções no Excel na Web com o ícone meus suplementos realçado](../images/excel-cf-online-register-add-in-1.png)

2. Escolha **Gerenciar Meus suplementos** e selecione **Carregar o Suplemento**.

3. Escolha **Procurar... ** e navegue até o diretório raiz do projeto criado pelo gerador Yeoman.

4. Selecione o arquivo **manifest. XML** e escolha **abrir**, escolha **Carregar**.

---

<ol start="5">
<li> Agora, vamos experimentar a nova função. Na célula <strong>B1</strong>, digite o texto <strong>= contoso. GETSTARCOUNT ("OfficeDev", "Excel-Custom-Functions")</strong> e pressione Enter. Você verá que o resultado na célula <strong>B1</strong> é o número atual de estrelas fornecido para o [repositório GitHub do Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions).</li>
</ol>

## <a name="create-a-streaming-asynchronous-custom-function"></a>Criar uma função personalizada assíncrona de streaming

A `getStarCount` função retorna o número de estrelas que um repositório tem em um momento específico no tempo. As funções personalizadas também podem retornar dados que estão sempre mudando. Essas funções são chamadas de fluxo de funções. Eles devem incluir um `invocation` parâmetro que se refira à célula de onde a função foi chamada. O `invocation` parâmetro é usado para atualizar o conteúdo da célula a qualquer momento.  

No exemplo de código a seguir, você verá que há duas funções `currentTime` e. `clock` A `currentTime` função é uma função estática que não usa streaming. Ele retorna a data como uma cadeia de caracteres. A `clock` função usa a `currentTime` função para fornecer a nova hora a cada segundo para uma célula no Excel. Ele usa `invocation.setResult` para entregar o tempo para a célula Excel e `invocation.onCanceled` para manipular o que ocorre quando a função é cancelada.

1. No projeto **starcount** , adicione o código a seguir a **./src/Functions/functions.js** e salve o arquivo.

```JS
/**
 * Returns the current time
 * @returns {string} String with the current time formatted for the current locale.
 */
function currentTime() {
  return new Date().toLocaleTimeString();
}

CustomFunctions.associate("CURRENTTIME", currentTime); 

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
CustomFunctions.associate("CLOCK", clock);
```

O `CustomFunctions.associate` código associa a `id` da função com o endereço de função da `CLOCK` em JavaScript para que o Excel possa ligar para a função.

2. Execute o seguinte comando para recriar o projeto.

    ```command&nbsp;line
    npm run build
    ```

3. Complete as etapas a seguir (para o Excel na Web ou Windows) para registrar novamente o suplemento no Excel. Você deve concluir estas etapas para que a nova função esteja disponível. 

# <a name="excel-on-windowstabexcel-windows"></a>[Excel no Windows](#tab/excel-windows)

1. Feche o Excel e abra novamente o Excel.

2. No Excel, escolha a guia **Inserir** e, em seguida, escolha a seta para baixo localizada à direita de **meus**suplementos.  ![Inserir faixa de opções no Excel no Windows com a seta meus suplementos realçada](../images/select-insert.png)

3. Na lista de suplementos disponíveis, encontre a seção suplementos do **desenvolvedor** e selecione o suplemento do **starcount** para registrá-lo.
    ![Inserir faixa de opções no Excel no Windows com o suplemento funções personalizadas do Excel realçado na lista meus suplementos](../images/list-starcount.png)

# <a name="excel-on-the-webtabexcel-online"></a>[Excel na Web](#tab/excel-online)

1. No Excel, escolha a guia **Inserir** e, em seguida, escolha **suplementos**.  ![Inserir faixa de opções no Excel na Web com o ícone meus suplementos realçado](../images/excel-cf-online-register-add-in-1.png)

2. Escolha **Gerenciar Meus suplementos** e selecione **Carregar o Suplemento**.

3. Escolha **Procurar... ** e navegue até o diretório raiz do projeto criado pelo gerador Yeoman.

4. Selecione o arquivo **manifest. XML** e escolha **abrir**, escolha **Carregar**.

--- 

<ol start="4">
<li>Agora, vamos experimentar a nova função. Na célula <strong>C1</strong>, digite o texto <strong>= contoso. RELÓGIO ())</strong> e pressione Enter. Você deve ver a data atual, que transmite uma atualização a cada segundo. Embora esse relógio seja apenas um cronômetro em um loop, você pode usar a mesma ideia de definir um timer em funções mais complexas que fazem solicitações da Web para dados em tempo real.</li>
</ol>

## <a name="next-steps"></a>Próximas etapas

Parabéns! Você criou um novo projeto de funções personalizadas, tentou uma função predefinida, criou uma função personalizada que solicita dados da Web e criou uma função personalizada que transmite dados. Você também pode experimentar a depuração dessa função usando [as instruções de depuração da função personalizada](../excel/custom-functions-debugging.md). Para saber mais sobre funções personalizadas no Excel, prossiga para o seguinte artigo:

> [!div class="nextstepaction"]
> [Criar funções personalizadas no Excel](../excel/custom-functions-overview.md)