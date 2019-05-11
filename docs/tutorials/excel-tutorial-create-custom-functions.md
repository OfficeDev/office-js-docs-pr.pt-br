---
title: Tutorial de funções personalizadas do Excel
description: Neste tutorial, você criará um suplemento do Excel que contém uma função personalizada que pode executar cálculos e solicitar ou transmitir dados da web.
ms.date: 05/08/2019
ms.prod: excel
ms.topic: tutorial
localization_priority: Normal
ms.openlocfilehash: ed9f16bdb330aa3f092e7d437ccfad6e056e07d4
ms.sourcegitcommit: a99be9c4771c45f3e07e781646e0e649aa47213f
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/11/2019
ms.locfileid: "33952191"
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

* Excel no Windows (versão de 64 bits 1810 ou posterior) ou Excel online

* Ingressar o [programa Office Insider](https://products.office.com/office-insider) (nível**Insider**, anteriormente chamado de "Insider – modo rápido")

## <a name="create-a-custom-functions-project"></a>Criar um projeto com funções personalizadas

 Para começar, você criará o projeto de código para criar o suplemento função personalizada. O [gerador Yeoman para suplementos do Office](https://www.npmjs.com/package/generator-office) configurará seu projeto com algumas funções personalizadas predefinidas que você pode experimentar. Se você já tiver executado o início rápido de funções personalizadas e gerado um projeto, continue a usar esse projeto e pule para [esta etapa](#create-a-custom-function-that-requests-data-from-the-web) .

1. Execute o comando a seguir e responda aos prompts da seguinte forma.
    
    ```command&nbsp;line
    yo office
    ```
    
    * **Escolha o tipo de projeto:** `Excel Custom Functions Add-in project (...)`
    * **Escolha o tipo de script:** `JavaScript`
    * **Qual será o nome do suplemento?** `stock-ticker`

    ![O gerador Yeoman para suplementos do Office solicita funções personalizadas](../images/yo-office-excel-cf.png)
    
    O gerador Yeoman criará os arquivos do projeto e instalará os componentes Node de suporte.

2. Navegue até a pasta raiz do projeto.
    
    ```command&nbsp;line
    cd stock-ticker
    ```

3. Crie um projeto.
    
    ```command&nbsp;line
    npm run build
    ```

4. Inicie o servidor local da web, que é executado no Node. Você pode experimentar o suplemento função personalizada no Excel no Windows ou no Excel online.

# <a name="excel-on-windowstabexcel-windows"></a>[Excel no Windows](#tab/excel-windows)

Para testar seu suplemento no Excel no Windows, execute o seguinte comando. Quando você executar este comando, o servidor Web local será iniciado e o Excel no Windows será aberto com seu suplemento carregado.

```command&nbsp;line
npm run start:desktop
```

> [!NOTE]
> Os Suplementos do Office devem usar HTTPS, e não HTTP, mesmo durante o desenvolvimento. Se você for solicitado a instalar um certificado após executar `npm run start:desktop`, aceite a solicitação para instalar o certificado que o gerador do Yeoman fornecer.

# <a name="excel-onlinetabexcel-online"></a>[Excel Online](#tab/excel-online)

Para testar seu suplemento no Excel online, execute o seguinte comando. Quando você executa este comando, o servidor Web local iniciará.

```command&nbsp;line
npm run start:web
```

> [!NOTE]
> Os Suplementos do Office devem usar HTTPS, e não HTTP, mesmo durante o desenvolvimento. Se você for solicitado a instalar um certificado após executar `npm run start:web`, aceite a solicitação para instalar o certificado que o gerador do Yeoman fornecer.

Para usar seu suplemento de funções personalizadas, abra uma nova pasta de trabalho no Excel online. Nesta pasta de trabalho, conclua as seguintes etapas para Sideload seu suplemento.

1. No Excel Online, escolha a guia **inserir** pressione e, em seguida, escolha **suplementos**.

   ![Inserir faixa de opções no Excel online com o ícone meus suplementos realçado](../images/excel-cf-online-register-add-in-1.png)
   
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

Integração de dados da Web é uma ótima maneira de ampliar o Excel por meio de funções personalizadas. Em seguida, você criará uma função personalizada chamada `stockPrice` que recebe uma citação ações de uma Web API e retorna o resultado para a célula de uma planilha. Esta função personalizada usa IEX Trading API, que é gratuito e não requer autenticação.

1. No projeto de **Cotações de ações** , localize o arquivo **./src/Functions/functions.js** e abra-o no editor de código.

2. Em **funções. js**, localize a `increment` função e adicione o código a seguir após essa função.

    ```js
    /**
    * Fetches current stock price
    * @customfunction 
    * @param {string} ticker Stock symbol
    * @returns {number} The current stock price.
    */
    function stockPrice(ticker) {
        var url = "https://api.iextrading.com/1.0/stock/" + ticker + "/price";
        return fetch(url)
            .then(function(response) {
                return response.text();
            })
            .then(function(text) {
                return parseFloat(text);
            });

        // Note: in case of an error, the returned rejected Promise
        //    will be bubbled up to Excel to indicate an error.
    }
    CustomFunctions.associate("STOCKPRICE", stockPrice);
    ```

    O `CustomFunctions.associate` código associa a `id` da função com o endereço de função da `stockPrice` em JavaScript para que o Excel possa ligar para a função.

3. Execute o seguinte comando para recriar o projeto.

    ```command&nbsp;line
    npm run build
    ```

4. Complete as etapas a seguir (para o Excel no Windows ou o Excel online) para registrar novamente o suplemento no Excel. Você deve concluir estas etapas para que a nova função esteja disponível. 

# <a name="excel-on-windowstabexcel-windows"></a>[Excel no Windows](#tab/excel-windows)

1. Feche o Excel e abra novamente o Excel.

2. No Excel, escolha a guia **Inserir** e, em seguida, escolha a seta para baixo localizada à direita de **meus**suplementos.  ![Inserir faixa de opções no Excel no Windows com a seta meus suplementos realçada](../images/select-insert.png)

3. Na lista de suplementos disponíveis, localize a seção**Suplementos do desenvolvedor** e selecione o suplemento **cotações** para registrá-lo.
    ![Inserir faixa de opções no Excel no Windows com o suplemento funções personalizadas do Excel realçado na lista meus suplementos](../images/list-stock-ticker-red.png)

# <a name="excel-onlinetabexcel-online"></a>[Excel Online](#tab/excel-online)

1. No Excel Online, escolha a guia **Inserir** e, em seguida, escolha **Suplementos**.  ![Insira a faixa de opções no Excel Online com o ícone Meus Suplementos realçado](../images/excel-cf-online-register-add-in-1.png)

2. Escolha **Gerenciar Meus suplementos** e selecione **Carregar o Suplemento**. 

3. Escolha **Procurar... ** e navegue até o diretório raiz do projeto criado pelo gerador Yeoman. 

4. Selecione o arquivo **manifest. XML** e escolha **abrir**, escolha **Carregar**.

---

<ol start="5">
<li> Agora, vamos experimentar a nova função. Na célula <strong>B1</strong>, digite o texto <strong>= da CONTOSO. STOCKPRICE("msft")</strong> e pressione enter. Você verá que o resultado na célula <strong>B1</strong> é o preço atual das ações para uma ação da Microsoft.</li>
</ol>

## <a name="create-a-streaming-asynchronous-custom-function"></a>Criar uma função personalizada assíncrona de streaming

A `stockPrice` função que você acabou de criar retorna o preço de uma ação em um momento específico, mas os preços das ações estão sempre mudando. Em seguida, você criará uma função personalizada chamada `stockPriceStream` esse é o preço de uma ação a cada 1000 milissegundos.

1. No projeto de **Cotações de ações** , adicione o seguinte código ao **/src/Functions/functions.js** e salve o arquivo.

    ```js
    /**
    * Streams real time stock price
    * @customfunction 
    * @param {string} ticker Stock symbol
    * @param {CustomFunctions.StreamingInvocation<number>} invocation
    */
    function stockPriceStream(ticker, invocation) {
        var updateFrequency = 1000 /* milliseconds*/;
        var isPending = false;

        var timer = setInterval(function() {
            // If there is already a pending request, skip this iteration:
            if (isPending) {
                return;
            }

            var url = "https://api.iextrading.com/1.0/stock/" + ticker + "/price";
            isPending = true;

            fetch(url)
                .then(function(response) {
                    return response.text();
                })
                .then(function(text) {
                    invocation.setResult(parseFloat(text));
                })
                .catch(function(error) {
                    invocation.setResult(error);
                })
                .then(function() {
                    isPending = false;
                });
        }, updateFrequency);

        invocation.onCanceled = () => {
            clearInterval(timer);
        };
    }
    CustomFunctions.associate("STOCKPRICESTREAM", stockPriceStream);
    ```
    
    O `CustomFunctions.associate` código associa a `id` da função com o endereço de função da `stockPriceStream` em JavaScript para que o Excel possa ligar para a função.
    
2. Execute o seguinte comando para recriar o projeto.

    ```command&nbsp;line
    npm run build
    ```

3. Complete as etapas a seguir (para o Excel no Windows ou o Excel online) para registrar novamente o suplemento no Excel. Você deve concluir estas etapas para que a nova função esteja disponível. 

# <a name="excel-on-windowstabexcel-windows"></a>[Excel no Windows](#tab/excel-windows)

1. Feche o Excel e abra novamente o Excel.

2. No Excel, escolha a guia **Inserir** e, em seguida, escolha a seta para baixo localizada à direita de **meus**suplementos.  ![Inserir faixa de opções no Excel no Windows com a seta meus suplementos realçada](../images/select-insert.png)

3. Na lista de suplementos disponíveis, localize a seção**Suplementos do desenvolvedor** e selecione o suplemento **cotações** para registrá-lo.
    ![Inserir faixa de opções no Excel no Windows com o suplemento funções personalizadas do Excel realçado na lista meus suplementos](../images/list-stock-ticker-red.png)

# <a name="excel-onlinetabexcel-online"></a>[Excel Online](#tab/excel-online)

1. No Excel Online, escolha a guia **Inserir** e, em seguida, escolha **Suplementos**.  ![Insira a faixa de opções no Excel Online com o ícone Meus Suplementos realçado](../images/excel-cf-online-register-add-in-1.png)

2. Escolha **Gerenciar Meus suplementos** e selecione **Carregar o Suplemento**.

3. Escolha **Procurar... ** e navegue até o diretório raiz do projeto criado pelo gerador Yeoman.

4. Selecione o arquivo **manifest. XML** e escolha **abrir**, escolha **Carregar**.

--- 

<ol start="4">
<li>Agora, vamos experimentar a nova função. Na célula <strong>C1</strong>, digite o texto <strong>= da CONTOSO. STOCKPRICESTREAM("msft")</strong> e pressione enter. Desde que o mercado de ações esteja aberto, você verá que o resultado na célula <strong>C1</strong> é constantemente atualizado para refletir o preço em tempo uma ação das ações da Microsoft.</li>
</ol>

## <a name="next-steps"></a>Próximas etapas

Parabéns! Neste tutorial, você criou um novo projeto de funções personalizadas, experimentou uma função predefinida, criou uma função personalizada que solicita dados da Web e criou uma função personalizada que transmite dados em tempo real da Web. Você também pode experimentar a depuração dessa função usando [as instruções de depuração da função personalizada](../excel/custom-functions-debugging.md). Para saber mais sobre funções personalizadas no Excel, prossiga para o seguinte artigo:

> [!div class="nextstepaction"]
> [Criar funções personalizadas no Excel](../excel/custom-functions-overview.md)

### <a name="legal-information"></a>Informações legais

Dados gratuito fornecidos pela [IEX](https://iextrading.com/developer/). Modo de exibição [termos de uso IEX](https://iextrading.com/api-exhibit-a/). O uso da Microsoft dA API IEX neste tutorial é apenas para fins educacionais.
