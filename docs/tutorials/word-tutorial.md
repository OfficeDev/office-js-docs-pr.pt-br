---
title: Tutorial de suplemento do Word
description: Neste tutorial, voc? criar? um suplemento do Word que insere (e substitui) intervalos de texto, par?grafos, imagens, HTML, tabelas e controles de conte?do. Você também aprenderá como formatar texto e como inserir (e substituir) conteúdo nos controles de conteúdo.
ms.date: 06/20/2019
ms.prod: word
ms.topic: tutorial
localization_priority: Normal
ms.openlocfilehash: 60397eb4afce60a0880f19be8296ad5fdce315a8
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771867"
---
# <a name="tutorial-create-a-word-task-pane-add-in"></a>Tutorial: Criar Suplemento do Painel de Tarefas no Word

Neste tutorial: você criará um suplemento do painel de tarefas no Word:

> [!div class="checklist"]
> * Insere um intervalo de texto
> * Formatos de texto
> * Substitui e insere texto em vários locais
> * Insere imagens, HTML e tabelas
> * Cria e atualiza os controles de conteúdo 

## <a name="prerequisites"></a>Pré-requisitos

Para usar este tutorial, você precisa instalar o seguinte.

- Word 2016, versão 1711 (build 8730.1000 do Clique para Executar) ou posterior. Talvez você precise ser um participante do programa Office Insider para ter essa versão. Para saber mais, confira a página [Seja um Office Insider](https://products.office.com/office-insider?tab=tab-1).

- [Nó](https://nodejs.org/en/) 

- [Git Bash](https://git-scm.com/downloads) (ou outro cliente Git)

## <a name="create-your-add-in-project"></a>Criar seu projeto do suplemento

Conclua as etapas a seguir para criar o projeto de suplemento do Word que você vai usar como base para este tutorial.

1. Clone o repositório do GitHub com o [Tutorial de suplemento do Word](https://github.com/OfficeDev/Word-Add-in-Tutorial).

2. Abra uma janela Git bash ou um prompt de sistema habilitado para Node.JS e navegue para a pasta **Iniciar** do projeto.

3. Execute o comando `npm install` para instalar as ferramentas e bibliotecas listadas no arquivo package.json. 

4. Execute as etapas de [instalação do certificado autoassinado](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) para confiar no certificado para o sistema operacional do seu computador de desenvolvimento.

## <a name="insert-a-range-of-text"></a>Inserir um intervalo de texto

Nesta etapa do tutorial, você testará programaticamente se o suplemento oferece suporte à versão atual do Word do usuário e inserirá um parágrafo no documento.

### <a name="code-the-add-in"></a>Codificação do suplemento

1. Abra o projeto em seu editor de código.

2. Abra o arquivo index.html.

3. Substitua `TODO1` pela marcação a seguir:

    ```html
    <button class="ms-Button" id="insert-paragraph">Insert Paragraph</button>
    ```

4. Abra o arquivo app.js.

5. Substitua o `TODO1` pelo código a seguir. O código determina se a versão do Word do usuário suporta uma versão do Word.js que inclui todas as APIs usadas em todos os estágios deste tutorial.dae Em um suplemento de produção, use o corpo do bloco condicional para ocultar ou desabilitar a interface do usuário que chame a APIs sem suporte. Isso permitirá que o usuário ainda use as partes do suplemento às quais a versão do Word dá suporte.

    ```js
    if (!Office.context.requirements.isSetSupported('WordApi', '1.3')) {
        console.log('Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.');
    }
    ```

6. Substitua o `TODO2` pelo código a seguir:

    ```js
    $('#insert-paragraph').click(insertParagraph);
    ```

7. Substitua o `TODO3` pelo código a seguir. Observação:

   - A lógica de negócios de Word.js será adicionada à função que passar por `Word.run`. Essa lógica não é executada imediatamente. Em vez disso, ela é adicionada à fila de comandos pendentes.

   - O método `context.sync` envia todos os comandos da fila para execução no Word.

   - O `Word.run` é seguido por um bloco `catch`. Essa é uma prática recomendada que você sempre deve seguir. 

    ```js
    function insertParagraph() {
        Word.run(function (context) {

            // TODO4: Queue commands to insert a paragraph into the document.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

8. Substitua `TODO4` pelo código a seguir. Observação:

   - O primeiro parâmetro para o método `insertParagraph` é o texto para o novo parágrafo.

   - O segundo parâmetro é o local dentro do corpo onde o parágrafo será inserido. Outras opções para inserir parágrafo, quando o objeto pai é o corpo, são "End" e "Replace".

    ```js
    var docBody = context.document.body;
    docBody.insertParagraph("Office has several versions, including Office 2016, Office 365 Click-to-Run, and Office on the web.",
                            "Start");
    ```

### <a name="test-the-add-in"></a>Testar o suplemento

1. Abra uma janela Git bash ou um prompt de sistema habilitado para Node.JS e navegue para a pasta **Iniciar** do projeto.

2. Execute o comando `npm run build` para transcompilar seu código-fonte ES6 para uma versão anterior do JavaScript com suporte de todos os hosts nos quais os suplementos do Office podem ser executados.

3. Execute o comando `npm start` para iniciar um servidor Web em um localhost.

4. Realize o sideload do suplemento usando um dos métodos a seguir:

    - Windows: [Realizar sideload de Suplementos do Office no Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)

    - Navegador da Web: [Sideload suplementos do Office no Office na Web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)

    - iPad e Mac: [Realizar sideload dos Suplementos do Office no iPad e Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)

5. No menu **Página Inicial** do Word, selecione **Mostrar Painel de Tarefas**.

6. No painel de tarefas, escolha **Inserir Parágrafo**.

7. Faça uma alteração no parágrafo.

8. Escolha novamente **Inserir Parágrafo**. O novo parágrafo está acima do anterior porque o método `insertParagraph` está inserido no início do corpo do documento.

    ![Tutorial do Word: Inserir Parágrafo](../images/word-tutorial-insert-paragraph.png)

## <a name="format-text"></a>Formatar texto

Nesta etapa do tutorial, você aplicará um estilo interno ao texto, aplicará um estilo personalizado ao texto e alterará a fonte do texto.

### <a name="apply-a-built-in-style-to-text"></a>Aplicar um estilo interno ao texto

1. Abra o projeto em seu editor de código. 

2. Abra o arquivo index.html.

3. Abaixo do `div`, que contém o botão `insert-paragraph`, adicione a marcação a seguir:

    ```html
    <div class="padding">            
        <button class="ms-Button" id="apply-style">Apply Style</button>            
    </div>
    ```

4. Abra o arquivo app.js.

5. Logo abaixo da linha que atribui um identificador de clique ao botão `insert-paragraph`, adicione o seguinte código:

    ```js
    $('#apply-style').click(applyStyle);
    ```

6. Logo abaixo da função `insertParagraph`, adicione a função a seguir:

    ```js
    function applyStyle() {
        Word.run(function (context) {
            
            // TODO1: Queue commands to style text.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ``` 

7. Substitua `TODO1` pelo código a seguir. O código aplica um estilo a um parágrafo, mas também é possível aplicar estilos em intervalos de texto.

    ```js
    var firstParagraph = context.document.body.paragraphs.getFirst();
    firstParagraph.styleBuiltIn = Word.Style.intenseReference;
    ``` 

### <a name="apply-a-custom-style-to-text"></a>Aplicar um estilo personalizado ao texto

1. Abra o arquivo index.html.

2. Abaixo do `div` que contém o botão `apply-style`, adicione a marcação a seguir:

    ```html
    <div class="padding">            
        <button class="ms-Button" id="apply-custom-style">Apply Custom Style</button>            
    </div>
    ```

3. Abra o arquivo app.js.

4. Abaixo da linha que atribui um identificador de clique ao botão `apply-style`, adicione o seguinte código:

    ```js
    $('#apply-custom-style').click(applyCustomStyle);
    ```

5. Abaixo da função `applyStyle`, adicione a função a seguir:

    ```js
    function applyCustomStyle() {
        Word.run(function (context) {
            
            // TODO1: Queue commands to apply the custom style.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ``` 

6. Substitua `TODO1` pelo código a seguir. O código aplica um estilo personalizado que ainda não existe. Você criará um estilo com o nome **MyCustomStyle** na etapa [Testar o suplemento](#test-the-add-in).

    ```js
    var lastParagraph = context.document.body.paragraphs.getLast();
    lastParagraph.style = "MyCustomStyle";
    ``` 

### <a name="change-the-font-of-text"></a>Alterar a fonte do texto

1. Abra o arquivo index.html.

2. Abaixo do `div` que contém o botão `apply-custom-style`, adicione a marcação a seguir:

    ```html
    <div class="padding">            
        <button class="ms-Button" id="change-font">Change Font</button>            
    </div>
    ```

3. Abra o arquivo app.js.

4. Abaixo da linha que atribui um identificador de clique ao botão `apply-custom-style`, adicione o seguinte código:

    ```js
    $('#change-font').click(changeFont);
    ```

5. Abaixo da função `applyCustomStyle`, adicione a função a seguir:

    ```js
    function changeFont() {
        Word.run(function (context) {
            
            // TODO1: Queue commands to apply a different font.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ``` 

6. Substitua `TODO1` pelo código a seguir. O código recebe uma referência para o segundo parágrafo usando o método `ParagraphCollection.getFirst` encadeado para o método `Paragraph.getNext`.

    ```js
    var secondParagraph = context.document.body.paragraphs.getFirst().getNext();
    secondParagraph.font.set({
            name: "Courier New",
            bold: true,
            size: 18
        });
    ``` 

### <a name="test-the-add-in"></a>Testar o suplemento

1. Na janela Git bash ou o prompt de sistema habilitado para Node.JS do tutorial anterior ainda estão abertos, digite Ctrl + C duas vezes para interromper a execução do servidor Web. Caso contrário, abra uma janela Git bash ou um prompt de sistema habilitado para Node.JS e navegue até a pasta **Iniciar** do projeto.

     > [!NOTE]
     > Embora o servidor de sincronização do navegador recarregue o suplemento no painel de tarefas sempre que você fizer uma alteração em algum arquivo, incluindo o arquivo app.js, ele não transcompila o JavaScript, portanto, é necessário repetir o comando de compilação para que as alterações em app.js as entrem em vigor. Para fazer isso, interrompa o processo do servidor para que o prompt apareça e você possa inserir o comando de compilação. Após a compilação, reinicie o servidor. As próximas etapas executam esse processo.

2. Execute o comando `npm run build` para transcompilar seu código-fonte ES6 para uma versão anterior do JavaScript com suporte de todos os hosts nos quais os suplementos do Office podem ser executados.

3. Execute o comando `npm start` para iniciar um servidor Web em um localhost.   

4. Feche o painel de tarefas para recarregá-lo e, no menu **Início**, selecione **Mostrar Painel de Tarefas** para reabrir o suplemento.

5. Verifique se há pelo menos três parágrafos no documento. É possível escolher **Inserir Parágrafo** três vezes. *Verifique com atenção se não há um parágrafo em branco no final do documento. Se houver, exclua-o.*

6. No Word, crie um estilo personalizado chamado "MyCustomStyle". Pode ter a formatação que você quiser.

7. Escolha o botão **Aplicar Estilo**. O primeiro parágrafo receberá o estilo interno **Referência Intensa**.

8. Escolha o botão **Aplicar Estilo Personalizado**. O último parágrafo receberá seu estilo personalizado. (Se parecer que nada acontece, talvez o último parágrafo esteja em branco. Se estiver, adicione um texto a ele).

9. Escolha o botão **Alterar Fonte**. A fonte do segundo parágrafo muda para 18 pt, negrito, Courier New.

    ![Tutorial do Word: Aplicar estilos e fonte](../images/word-tutorial-apply-styles-and-font.png)

## <a name="replace-text-and-insert-text"></a>Substituir texto e inserir texto

Nesta etapa o tutorial, você adicionará texto dentro e fora dos intervalos de texto selecionados, e substituirá o texto de um intervalo selecionado.

### <a name="add-text-inside-a-range"></a>Adicionar texto dentro de um intervalo

1. Abra o projeto em seu editor de código.

2. Abra o arquivo index.html.

3. Abaixo do `div` que contém o botão `change-font`, adicione a marcação a seguir:

    ```html
    <div class="padding">
        <button class="ms-Button" id="insert-text-into-range">Insert Abbreviation</button>
    </div>
    ```

4. Abra o arquivo app.js.

5. Abaixo da linha que atribui um identificador de clique ao botão `change-font`, adicione o seguinte código:

    ```js
    $('#insert-text-into-range').click(insertTextIntoRange);
    ```

6. Abaixo da função `changeFont`, adicione a função a seguir:

    ```js
    function insertTextIntoRange() {
        Word.run(function (context) {

            // TODO1: Queue commands to insert text into a selected range.

            // TODO2: Load the text of the range and sync so that the
            //        current range text can be read.

            // TODO3: Queue commands to repeat the text of the original
            //        range at the end of the document.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ``` 

7. Substitua `TODO1` pelo código a seguir. Observação:

   - o método serve para inserir a abreviação ["(C2R)"] no final do Intervalo cujo texto é "Clique para Executar". Para simplificar, ele faz uma pressuposição de que a cadeia de caracteres está presente, e que o usuário a selecionou.

   - O primeiro parâmetro do método `Range.insertText` é a cadeia de caracteres a ser inserida no objeto `Range`.

   - O segundo parâmetro especifica onde no intervalo, o texto adicional deve ser inserido. Além de "Fim", as outras opções possíveis são "Início", "Antes", "Depois" e "Substituir". 

   - A diferença entre "Fim" e "Depois" é que "Fim" insere o novo texto dentro o final do intervalo existente, mas "Depois" cria um novo intervalo com a cadeia de caracteres e insere o novo intervalo após o intervalo existente. Da mesma forma, "Início" insere o texto dentro do início do intervalo existente, e "Antes" insere um novo intervalo. "Substituir" substitui o texto do intervalo existente pela cadeia de caracteres do primeiro parâmetro.

   - Você viu em um estágio anterior do tutorial que os métodos insert* do objeto de corpo não têm as opções "Antes" e "Depois". Isso ocorre porque não é possível colocar o conteúdo fora do corpo do documento.

    ```js
    var doc = context.document;
    var originalRange = doc.getSelection();
    originalRange.insertText(" (C2R)", "End");
    ```

8. Vamos deixar `TODO2` de lado até a próxima seção. Substitua `TODO3` pelo código a seguir. Esse código é semelhante ao código que você criou no primeiro estágio do tutorial, exceto que, agora, você está inserindo um novo parágrafo no final do documento, em vez de no início. Este novo parágrafo demonstrará que o novo texto agora faz parte do intervalo original.

    ```js
    doc.body.insertParagraph("Original range: " + originalRange.text, "End");
    ```

### <a name="add-code-to-fetch-document-properties-into-the-task-panes-script-objects"></a>Adicione código para buscar propriedades do documento em objetos de script do painel de tarefas

Em todas as funções anteriores desta série de tutoriais, você colocou em fila comandos para *gravar* no documento do Office. Cada função terminou com uma chamada para o método `context.sync()`, que envia os comandos em fila para o documento a ser executado. Entretanto, o código adicionado na última etapa chama a propriedade `originalRange.text` e essa é uma grande diferença das funções anteriores que você escreveu, pois o objeto `originalRange` é apenas um objeto de proxy que existe no script do seu painel de tarefas. Ele não sabe qual é o texto real do intervalo no documento, portanto, sua propriedade `text` não pode ter um valor real. Primeiro, é necessário buscar o valor de texto do intervalo no documento e usá-lo para definir o valor de `originalRange.text`. Somente então será possível chamar `originalRange.text` sem causar uma exceção. Esse processo de busca tem três etapas:

   1. Coloque em fila um comando para carregar (ou seja, fetch) as propriedades que seu código precisa ler.

   2. Chame o método `sync` do objeto de contexto para enviar o comando em fila para o documento para execução e retornar as informações solicitadas.

   3. Como o método `sync` é assíncrono, certifique-se de que ele tenha sido concluído antes que o código chame as propriedades que foram buscadas.

Estas etapas devem ser concluídas sempre que seu código precisar *ler* informações do documento do Office.

1. Substitua `TODO2` pelo código a seguir.
  
    ```js
    originalRange.load("text");
    return context.sync()
        .then(function() {

                // TODO4: Move the doc.body.insertParagraph line here.

            }
        )
            // TODO5: Move the final call of context.sync here and ensure
            //        that it does not run until the insertParagraph has
            //        been queued.
    ```

2. Você não pode ter duas instruções `return` no mesmo caminho de código sem ramificações, portanto, exclua a linha final `return context.sync();` no final de `Word.run`. Você adicionará um novo final `context.sync` posteriormente neste tutorial.

3. Recorte a linha `doc.body.insertParagraph` e cole no lugar de `TODO4`.

4. Substitua `TODO5` pelo código a seguir. Observação:

   - Passar o método `sync` para uma função `then` garante que ele não seja executado até que a lógica `insertParagraph` tenha sido enfileirada.

   - O método `then` invoca qualquer função que é passada para ele e não é recomendável que `sync` seja chamado duas vezes, portanto, remova os "()" do fim de context.sync.

    ```js
    .then(context.sync);
    ```

Quando terminar, a função inteira deve se parecer com o seguinte:

```js
function insertTextIntoRange() {
    Word.run(function (context) {

        var doc = context.document;
        var originalRange = doc.getSelection();
        originalRange.insertText(" (C2R)", "End");

        originalRange.load("text");
        return context.sync()
            .then(function() {
                        doc.body.insertParagraph("Current text of original range: " + originalRange.text,
                                                "End");
                }
            )
            .then(context.sync);
    })
    .catch(function (error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });
}
```

### <a name="add-text-between-ranges"></a>Adicionar texto entre intervalos

1. Abra o arquivo index.html.

2. Abaixo do `div` que contém o botão `insert-text-into-range`, adicione a marcação a seguir:

    ```html
    <div class="padding">
        <button class="ms-Button" id="insert-text-outside-range">Add Version Info</button>
    </div>
    ```

3. Abra o arquivo app.js.

4. Abaixo da linha que atribui um identificador de clique ao botão `insert-text-into-range`, adicione o seguinte código:

    ```js
    $('#insert-text-outside-range').click(insertTextBeforeRange);
    ```

5. Abaixo da função `insertTextIntoRange`, adicione a função a seguir:

    ```js
    function insertTextBeforeRange() {
        Word.run(function (context) {

            // TODO1: Queue commands to insert a new range before the
            //        selected range.

            // TODO2: Load the text of the original range and sync so that the
            //        range text can be read and inserted.

        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

6. Substitua `TODO1` pelo código a seguir. Observação:

   - O método serve para adicionar um intervalo cujo texto seja "Office 2019", antes do intervalo com o texto "Office 365". Para simplificar, ele faz uma pressuposição de que a cadeia de caracteres está presente, e que o usuário a selecionou.

   - O primeiro parâmetro do método `Range.insertText` é a cadeia de caracteres a ser adicionada.

   - O segundo parâmetro especifica onde no intervalo, o texto adicional deve ser inserido. Para ter mais detalhes sobre as opções de local, confira a discussão anterior sobre a função `insertTextIntoRange`.

    ```js
    var doc = context.document;
    var originalRange = doc.getSelection();
    originalRange.insertText("Office 2019, ", "Before");
    ```

7. Substitua `TODO2` pelo código a seguir.

     ```js
    originalRange.load("text");
    return context.sync()
        .then(function() {

                // TODO3: Queue commands to insert the original range as a
                //        paragraph at the end of the document.

                }
            )

            // TODO4: Make a final call of context.sync here and ensure
            //        that it does not run until the insertParagraph has
            //        been queued.
    ```

8. Substitua `TODO3` pelo código a seguir. Este novo parágrafo demonstrará que o novo texto ***não*** faz parte do intervalo original selecionado. O intervalo original ainda contém o texto que tinha quando foi selecionado.

    ```js
    doc.body.insertParagraph("Current text of original range: " + originalRange.text,
                             "End");
    ```

9. Substitua `TODO4` pelo código a seguir:

    ```js
    .then(context.sync);
    ```

### <a name="replace-the-text-of-a-range"></a>Substitua o texto de um intervalo.

1. Abra o arquivo index.html.

2. Abaixo do `div` que contém o botão `insert-text-outside-range`, adicione a marcação a seguir:

    ```html
    <div class="padding">
        <button class="ms-Button" id="replace-text">Change Quantity Term</button>
    </div>
    ```

3. Abra o arquivo app.js.

4. Abaixo da linha que atribui um identificador de clique ao botão `insert-text-outside-range`, adicione o seguinte código:

    ```js
    $('#replace-text').click(replaceText);
    ```

5. Abaixo da função `insertTextBeforeRange`, adicione a função a seguir:

    ```js
    function replaceText() {
        Word.run(function (context) {

            // TODO1: Queue commands to replace the text.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

6. Substitua `TODO1` pelo código a seguir. O método serve para substituir a cadeia de caracteres "várias" pela cadeia "muitos". Para simplificar, ele faz uma pressuposição de que a cadeia de caracteres está presente, e que o usuário a selecionou.

    ```js
    var doc = context.document;
    var originalRange = doc.getSelection();
    originalRange.insertText("many", "Replace");
    ```

### <a name="test-the-add-in"></a>Testar o suplemento

1. Se a janela Git bash ou o prompt de sistema habilitado para Node.JS do tutorial anterior ainda estiverem abertos, digite Ctrl + C duas vezes para interromper a execução do servidor Web. Caso contrário, abra uma janela Git bash ou um prompt de sistema habilitado para Node.JS e navegue até a pasta **Iniciar** do projeto.

     > [!NOTE]
     > Embora o servidor de sincronização do navegador recarregue o suplemento no painel de tarefas sempre que você fizer uma alteração em algum arquivo, incluindo o arquivo app.js, ele não transcompila o JavaScript, portanto, é necessário repetir o comando de compilação para que as alterações em app.js as entrem em vigor. Para fazer isso, interrompa o processo do servidor para que o prompt apareça e você possa inserir o comando de compilação. Após a compilação, reinicie o servidor. As próximas etapas executam esse processo.

2. Execute o comando `npm run build` para transcompilar seu código-fonte ES6 para uma versão anterior do JavaScript com suporte de todos os hosts nos quais os suplementos do Office podem ser executados.

3. Execute o comando `npm start` para iniciar um servidor Web em um localhost.

4. Feche o painel de tarefas para recarregá-lo e, no menu **Página Inicial**, selecione **Mostrar Painel de Tarefas** para reabrir o suplemento.

5. No painel de tarefas, escolha **Inserir Parágrafo** para garantir que haja um parágrafo no início do documento.

6. Selecione um texto. Selecionar a frase "Clique para Executar" fará mais sentido. *Tenha cuidado para não incluir o espaço anterior ou seguinte na seleção.*

7. Escolha o botão **Inserir Abreviação**. "(C2R)" é adicionado. Na parte inferior do documento, um novo parágrafo é adicionado com o texto inteiro expandido porque a nova cadeia de caracteres foi adicionada ao intervalo existente.

8. Selecione um texto. Selecionar a frase "Office 365" fará mais sentido. *Tenha cuidado para não incluir o espaço anterior ou seguinte na seleção.*

9. Escolha o botão **Adicionar Informações de Versão**. "Office 2019" está inserido entre "Office 2016" e "Office 365". Na parte inferior do documento um novo parágrafo foi adicionado, mas ele contém apenas o texto selecionado originalmente porque a nova cadeia de caracteres tornou-se um intervalo novo, em vez de ser adicionada ao intervalo original.

10. Selecione um texto. Selecionar a palavra "vários" fará mais sentido. *Tenha cuidado para não incluir o espaço anterior ou seguinte na seleção.*

11. Escolha o botão **Alterar Termo de Quantidade**. "muitos" substitui o texto selecionado.

    ![Tutorial do Word: texto adicionado e substituído](../images/word-tutorial-text-replace.png)

## <a name="insert-images-html-and-tables"></a>Inserir imagens, HTML e tabelas

Nesta etapa do tutorial, você aprenderá a inserir imagens, HTML e tabelas no documento.

### <a name="insert-an-image"></a>Inserir uma imagem

1. Abra o projeto em seu editor de código.

2. Abra o arquivo index.html.

3. Abaixo do `div` que contém o botão `replace-text`, adicione a marcação a seguir:

    ```html
    <div class="padding">
        <button class="ms-Button" id="insert-image">Insert Image</button>
    </div>
    ```

4. Abra o arquivo app.js.

5. Na parte superior do arquivo, logo abaixo da linha use-strict, adicione a seguinte linha. Essa linha importa uma variável de outro arquivo. A variável é uma cadeia de caracteres base 64 que codifica uma imagem. Para ver a cadeia de caracteres codificada, abra o arquivo base64Image.js na raiz do projeto.

    ```js
    import { base64Image } from "./base64Image";
    ```

6. Abaixo da linha que atribui um identificador de clique ao botão `replace-text`, adicione o seguinte código:

    ```js
    $('#insert-image').click(insertImage);
    ```

7. Abaixo da função `replaceText`, adicione a função a seguir:

    ```js
    function insertImage() {
        Word.run(function (context) {

            // TODO1: Queue commands to insert an image.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

8. Substitua `TODO1` pelo código a seguir. Esta linha insere a imagem codificada em base 64 no final do documento. (O objeto `Paragraph` também tem um método `insertInlinePictureFromBase64` e outros métodos `insert*`. Confira a seção insertHTML a seguir para conferir um exemplo).

    ```js
    context.document.body.insertInlinePictureFromBase64(base64Image, "End");
    ```

### <a name="insert-html"></a>Inserir HTML

1. Abra o arquivo index.html.

2. Abaixo do `div` que contém o botão `insert-image`, adicione a marcação a seguir:

    ```html
    <div class="padding">
        <button class="ms-Button" id="insert-html">Insert HTML</button>
    </div>
    ```

3. Abra o arquivo app.js.

4. Abaixo da linha que atribui um identificador de clique ao botão `insert-image`, adicione o seguinte código:

    ```js
    $('#insert-html').click(insertHTML);
    ```

5. Abaixo da função `insertImage`, adicione a função a seguir:

    ```js
    function insertHTML() {
        Word.run(function (context) {

            // TODO1: Queue commands to insert a string of HTML.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

6. Substitua `TODO1` pelo código a seguir. Observação:

   - A primeira linha adiciona um parágrafo em branco ao final do documento. 

   - A segunda linha insere uma cadeia de caracteres de HTML no final do parágrafo; especificamente dois parágrafos, um formatado com a fonte Verdana, e o outro com estilo padrão de documento do Word. (Conforme mostrado anteriormente no método `insertImage`, o objeto `context.document.body` também tem os métodos `insert*`).

    ```js
    var blankParagraph = context.document.body.paragraphs.getLast().insertParagraph("", "After");
    blankParagraph.insertHtml('<p style="font-family: verdana;">Inserted HTML.</p><p>Another paragraph</p>', "End");
    ```

### <a name="insert-a-table"></a>Inserir uma tabela

1. Abra o arquivo index.html.

2. Abaixo do `div` que contém o botão `insert-html`, adicione a marcação a seguir:

    ```html
    <div class="padding">
        <button class="ms-Button" id="insert-table">Insert Table</button>
    </div>
    ```

3. Abra o arquivo app.js.

4. Abaixo da linha que atribui um identificador de clique ao botão `insert-html`, adicione o seguinte código:

    ```js
    $('#insert-table').click(insertTable);
    ```

5. Abaixo da função `insertHTML`, adicione a função a seguir:

    ```js
    function insertTable() {
        Word.run(function (context) {

            // TODO1: Queue commands to get a reference to the paragraph
            //        that will proceed the table.

            // TODO2: Queue commands to create a table and populate it with data.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

6. Substitua `TODO1` pelo código a seguir. Essa linha usa o método `ParagraphCollection.getFirst` para obter uma referência do primeiro parágrafo e, depois, usa o método `Paragraph.getNext` para obter uma referência para o segundo parágrafo.

    ```js
    var secondParagraph = context.document.body.paragraphs.getFirst().getNext();
    ```

7. Substitua `TODO2` pelo código a seguir. Observação:

   - Os dois primeiros parâmetros do método `insertTable` especificam o número de linhas e colunas.

   - O terceiro parâmetro especifica onde inserir a tabela, nesse caso, depois do parágrafo.

   - O quarto parâmetro é uma matriz bidimensional que define os valores das células da tabela.

   - A tabela terá um estilo padrão simples, mas o método `insertTable` retornará um objeto `Table` com muitos membros, e alguns deles são usados para alterar o estilo de tabela.

    ```js
    var tableData = [
            ["Name", "ID", "Birth City"],
            ["Bob", "434", "Chicago"],
            ["Sue", "719", "Havana"],
        ];
    secondParagraph.insertTable(3, 3, "After", tableData);
    ```

### <a name="test-the-add-in"></a>Testar o suplemento

1. Se a janela Git bash ou o prompt de sistema habilitado para Node.JS do tutorial anterior ainda estiverem abertos, digite Ctrl+C duas vezes para interromper a execução do servidor Web. Caso contrário, abra uma janela Git bash ou um prompt de sistema habilitado para Node.JS e navegue até a pasta **Iniciar** do projeto.

     > [!NOTE]
     > Embora o servidor de sincronização do navegador recarregue o suplemento no painel de tarefas sempre que você fizer uma alteração em algum arquivo, incluindo o arquivo app.js, ele não transcompila o JavaScript, portanto, é necessário repetir o comando de compilação para que as alterações em app.js as entrem em vigor. Para fazer isso, interrompa o processo do servidor para que o prompt apareça e você possa inserir o comando de compilação. Após a compilação, reinicie o servidor. As próximas etapas executam esse processo.

2. Execute o comando `npm run build` para transcompilar seu código-fonte ES6 para uma versão anterior do JavaScript com suporte de todos os hosts nos quais os suplementos do Office podem ser executados.

3. Execute o comando `npm start` para iniciar um servidor Web em um localhost.

4. Feche o painel de tarefas para recarregá-lo e, no menu **Página Inicial**, selecione **Mostrar Painel de Tarefas** para reabrir o suplemento.

5. No painel de tarefas, escolha **Inserir Parágrafo** pelo menos três vezes para garantir que haja alguns parágrafos no documento.

6. Escolha o botão **Inserir Imagem**. Uma imagem é inserida no final do documento.

7. Escolha o botão **Inserir HTML**. Dois parágrafos são inseridos no final do documento, e o primeiro tem a fonte Verdana.

8. Escolha o botão **Inserir Tabela**. Uma tabela é inserida após o segundo parágrafo.

    ![Tutorial do Word: Inserir imagem, HTML e tabela](../images/word-tutorial-insert-image-html-table.png)

## <a name="create-and-update-content-controls"></a>Criar e atualizar os controles de conteúdo

Nesta etapa do tutorial, você aprenderá a criar controles de conteúdo de Rich Text no documento e, depois, como inserir e substituir conteúdo nos controles.

> [!NOTE]
> Há vários tipos de controles de conteúdo que podem ser adicionados a um documento do Word por meio da interface do usuário. Porém, no momento, só há suporte para controles de conteúdo de Rich Text no Word.js.
>
> Antes de começar esta etapa do tutorial, recomendamos a criação e manipulação dos controles de conteúdo de Rich Text por meio da interface do usuário do Word, para se familiarizar com os controles e suas propriedades. Para saber mais detalhes, confira [Criar formulários para preenchimento ou impressão no Word](https://support.office.com/article/create-forms-that-users-complete-or-print-in-word-040c5cc1-e309-445b-94ac-542f732c8c8b).

### <a name="create-a-content-control"></a>Criar um controle de conteúdo

1. Abra o projeto em seu editor de código.

2. Abra o arquivo index.html.

3. Abaixo do `div` que contém o botão `replace-text`, adicione a marcação a seguir:

    ```html
    <div class="padding">
        <button class="ms-Button" id="create-content-control">Create Content Control</button>
    </div>
    ```

4. Abra o arquivo app.js.

5. Abaixo da linha que atribui um identificador de clique ao botão `insert-table`, adicione o seguinte código:

    ```js
    $('#create-content-control').click(createContentControl);
    ```

6. Abaixo da função `insertTable`, adicione a função a seguir:

    ```js
    function createContentControl() {
        Word.run(function (context) {

            // TODO1: Queue commands to create a content control.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

7. Substitua `TODO1` pelo código a seguir. Observação:

   - o código tem como objetivo dispor a frase "Office 365" em um controle de conteúdo. Para simplificar, ele faz uma pressuposição de que a cadeia de caracteres está presente, e que o usuário a selecionou.

   - A propriedade `ContentControl.title` especifica o título visível do controle de conteúdo.

   - A propriedade `ContentControl.tag` especifica uma marca que pode ser usada para obter uma referência a um controle de conteúdo usando o método `ContentControlCollection.getByTag`, que você usará em uma função posterior.

   - A propriedade `ContentControl.appearance` especifica a aparência do controle. Usar o valor "Tags" significa que o controle será encapsulado entre marcas de abertura e fechamento, e a marca de abertura terá o título do controle de conteúdo. Outros valores possíveis são "BoundingBox" e "None".

   - A propriedade `ContentControl.color` especifica a cor das marcas ou da borda da caixa delimitadora.

    ```js
    var serviceNameRange = context.document.getSelection();
    var serviceNameContentControl = serviceNameRange.insertContentControl();
    serviceNameContentControl.title = "Service Name";
    serviceNameContentControl.tag = "serviceName";
    serviceNameContentControl.appearance = "Tags";
    serviceNameContentControl.color = "blue";
    ```

### <a name="replace-the-content-of-the-content-control"></a>Substituir o conteúdo do controle de conteúdo

1. Abra o arquivo index.html.

2. Abaixo do `div` que contém o botão `create-content-control`, adicione a marcação a seguir:

    ```html
    <div class="padding">
        <button class="ms-Button" id="replace-content-in-control">Rename Service</button>
    </div>
    ```

3. Abra o arquivo app.js.

4. Abaixo da linha que atribui um identificador de clique ao botão `create-content-control`, adicione o seguinte código:

    ```js
    $('#replace-content-in-control').click(replaceContentInControl);
    ```

5. Abaixo da função `createContentControl`, adicione a função a seguir:

    ```js
    function replaceContentInControl() {
        Word.run(function (context) {

            // TODO1: Queue commands to replace the text in the Service Name
            //        content control.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

6. Substitua `TODO1` pelo código a seguir. Observação:

    - O método `ContentControlCollection.getByTag` retorna um `ContentControlCollection` de todos os controles de conteúdo da marca especificada. Usamos `getFirst` para obter uma referência do controle desejado.

    ```js
    var serviceNameContentControl = context.document.contentControls.getByTag("serviceName").getFirst();
    serviceNameContentControl.insertText("Fabrikam Online Productivity Suite", "Replace");
    ```

### <a name="test-the-add-in"></a>Testar o suplemento

1. Se a janela Git bash ou o prompt de sistema habilitado para Node.JS do tutorial anterior ainda estiverem abertos, digite Ctrl+C duas vezes para interromper a execução do servidor Web. Caso contrário, abra uma janela Git bash ou um prompt de sistema habilitado para Node.JS e navegue até a pasta **Iniciar** do projeto.

     > [!NOTE]
     > Embora o servidor de sincronização do navegador recarregue o suplemento no painel de tarefas sempre que você fizer uma alteração em algum arquivo, incluindo o arquivo app.js, ele não transcompila o JavaScript, portanto, é necessário repetir o comando de compilação para que as alterações em app.js as entrem em vigor. Para fazer isso, interrompa o processo do servidor para que o prompt apareça e você possa inserir o comando de compilação. Após a compilação, reinicie o servidor. As próximas etapas executam esse processo.

2. Execute o comando `npm run build` para transcompilar seu código-fonte ES6 para uma versão anterior do JavaScript com suporte de todos os hosts nos quais os suplementos do Office podem ser executados.

3. Execute o comando `npm start` para iniciar um servidor Web em um localhost.

4. Feche o painel de tarefas para recarregá-lo e, no menu **Página Inicial**, selecione **Mostrar Painel de Tarefas** para reabrir o suplemento.

5. No painel de tarefas, escolha **Inserir Parágrafo** para garantir que haja um parágrafo com "Office 365" no início do documento.

6. Selecione a frase "Office 365" no parágrafo que você adicionou e escolha o botão **Criar Controle de Conteúdo**. A frase está envolvida por marcas chamadas "Nome do Serviço".

7. Escolha o botão **Renomear Serviço**. O texto do controle de conteúdo muda para "Fabrikam Online Productivity Suite".

    ![Tutorial do Word - Criar o controle de conteúdo e alterar seu texto](../images/word-tutorial-content-control.png)

## <a name="next-steps"></a>Próximas etapas

Neste tutorial, você criou um suplemento do painel de tarefas do Word que insere e substitui texto, imagens e outro conteúdo em um documento do Word. Para saber mais sobre o desenvolvimento de suplementos do Word, continue no seguinte artigo:

> [!div class="nextstepaction"]
> [Visão geral dos suplementos do Word](../word/word-add-ins-programming-overview.md)
