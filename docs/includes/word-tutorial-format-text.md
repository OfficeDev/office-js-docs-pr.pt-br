Nesta etapa do tutorial, você mudará a fonte do texto e usará estilos internos e personalizados no texto.

> [!NOTE]
> Esta página descreve uma etapa individual de um tutorial de suplemento do Word. Se você chegou aqui por meio dos resultados de mecanismos de pesquisa ou por outro link direto, acesse a página de introdução do [tutorial de suplemento do Word](../tutorials/word-tutorial.yml) para começar pelo início.

## <a name="apply-a-built-in-style-to-text"></a>Aplicar um estilo interno ao texto

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
    const firstParagraph = context.document.body.paragraphs.getFirst();
    firstParagraph.styleBuiltIn = Word.Style.intenseReference;
    ``` 

## <a name="apply-a-custom-style-to-text"></a>Aplicar um estilo personalizado ao texto

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

7. Substitua `TODO1` pelo código a seguir. O código aplica um estilo personalizado que ainda não existe. Você criará um estilo com o nome **MyCustomStyle** na etapa [Testar o suplemento](#test-the-add-in).

    ```js
    const lastParagraph = context.document.body.paragraphs.getLast();
    lastParagraph.style = "MyCustomStyle";
    ``` 

## <a name="change-the-font-of-text"></a>Alterar a fonte do texto

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

7. Substitua `TODO1` pelo código a seguir. O código recebe uma referência para o segundo parágrafo usando o método `ParagraphCollection.getFirst` encadeado para o método `Paragraph.getNext`.

    ```js
    const secondParagraph = context.document.body.paragraphs.getFirst().getNext();
    secondParagraph.font.set({
            name: "Courier New",
            bold: true,
            size: 18
        });
    ``` 

## <a name="test-the-add-in"></a>Testar o suplemento

1. Se a janela Git bash ou o prompt de sistema habilitado para Node.JS do tutorial anterior ainda estiverem abertos, digite Ctrl+C duas vezes para interromper a execução do servidor Web. Caso contrário, abra uma janela Git bash ou um prompt de sistema habilitado para Node.JS e navegue até a pasta **Iniciar** do projeto.

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
