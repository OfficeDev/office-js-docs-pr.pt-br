Nesta etapa do tutorial, você testará programaticamente se o suplemento oferece suporte à versão atual do Word do usuário e inserirá um parágrafo no documento.

> [!NOTE]
> Esta página descreve uma etapa individual de um tutorial de suplemento do Word. Se você chegou aqui por meio dos resultados de mecanismos de pesquisa ou por outro link direto, acesse a página de introdução do [tutorial de suplemento do Word](../tutorials/word-tutorial.yml) para começar pelo início.

## <a name="code-the-add-in"></a>Codificação do suplemento

1. Abra o projeto em seu editor de código. 
2. Abra o arquivo index.html.
3. Substitua `TODO1` pela marcação a seguir:

    ```html
    <button class="ms-Button" id="insert-paragraph">Insert Paragraph</button>
    ```

4. Abra o arquivo app.js.
5. Substitua o `TODO1` pelo código a seguir. O código determina se a versão do Word do usuário suporta uma versão do Word.js que inclui todas as APIs usadas em todos os estágios deste tutorial.dae Em um suplemento de produção, use o corpo do bloco condicional para ocultar ou desabilitar a interface do usuário que chame a APIs sem suporte. Isso permitirá que o usuário ainda use as partes do suplemento às quais a versão do Word dá suporte.

    ```js
    if (!Office.context.requirements.isSetSupported('WordApi', 1.3)) {
        console.log('Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.');
    } 
    ```

6. Substitua o `TODO2` pelo código a seguir:

    ```js
    $('#insert-paragraph').click(insertParagraph);
    ```

7. Substitua o `TODO3` pelo código a seguir. Observe o seguinte:
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
    const docBody = context.document.body;
    docBody.insertParagraph("Office has several versions, including Office 2016, Office 365 Click-to-Run, and Office Online.",
                            "Start");   
    ``` 

## <a name="test-the-add-in"></a>Testar o suplemento

1. Abra uma janela Git bash ou um prompt de sistema habilitado para Node.JS e navegue para a pasta **Iniciar** do projeto.
2. Execute o comando `npm run build` para transcompilar seu código-fonte ES6 para uma versão anterior do JavaScript com suporte de todos os hosts nos quais os suplementos do Office podem ser executados.
3. Execute o comando `npm start` para iniciar um servidor Web em um localhost.   
4. Realize o sideload do suplemento usando um dos métodos a seguir:
    - Windows: [Realizar sideload de Suplementos do Office no Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
    - Word Online: [Realizar sideload dos Suplementos do Office no Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)
    - iPad e Mac: [Realizar sideload dos Suplementos do Office no iPad e Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)
5. No menu **Página Inicial** do Word, selecione **Mostrar Painel de Tarefas**.
6. No painel de tarefas, escolha **Inserir Parágrafo**.
7. Faça uma alteração no parágrafo. 
8. Escolha novamente **Inserir Parágrafo**. O novo parágrafo está acima do anterior porque o método `insertParagraph` está inserido no "início" do corpo do documento.

    ![Tutorial do Word: Inserir Parágrafo](../images/word-tutorial-insert-paragraph.png)
