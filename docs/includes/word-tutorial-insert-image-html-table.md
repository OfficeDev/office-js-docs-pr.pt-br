Nesta etapa do tutorial, você aprenderá a inserir imagens, HTML e tabelas no documento.

> [!NOTE]
> Esta página descreve uma etapa individual de um tutorial de suplemento do Word. Se você chegou aqui por meio dos resultados de mecanismos de pesquisa ou por outro link direto, acesse a página de introdução do [tutorial de suplemento do Word](../tutorials/word-tutorial.yml) para começar pelo início.

## <a name="insert-an-image"></a>Inserir uma imagem

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

## <a name="insert-html"></a>Inserir HTML

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
    const blankParagraph = context.document.body.paragraphs.getLast().insertParagraph("", "After");
    blankParagraph.insertHtml('<p style="font-family: verdana;">Inserted HTML.</p><p>Another paragraph</p>', "End");
    ```

## <a name="insert-table"></a>Inserir Tabela

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
    const secondParagraph = context.document.body.paragraphs.getFirst().getNext();
    ```

7. Substitua `TODO2` pelo código a seguir. Observação:
   - Os dois primeiros parâmetros do método `insertTable` especificam o número de linhas e colunas.
   - O terceiro parâmetro especifica onde inserir a tabela, nesse caso, depois do parágrafo.
   - O quarto parâmetro é uma matriz bidimensional que define os valores das células da tabela.
   - A tabela terá um estilo padrão simples, mas o método `insertTable` retornará um objeto `Table` com muitos membros, e alguns deles são usados para alterar o estilo de tabela.

    ```js
    const tableData = [
            ["Name", "ID", "Birth City"],
            ["Bob", "434", "Chicago"],
            ["Sue", "719", "Havana"],
        ];
    secondParagraph.insertTable(3, 3, "After", tableData);
    ```

## <a name="test-the-add-in"></a>Testar o suplemento


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
