Nesta etapa do tutorial, voc? aprender? a criar controles de conte?do de Rich Text no documento e, depois, como inserir e substituir conte?do nos controles. 

> [!NOTE]
> Esta p?gina descreve uma etapa individual de um tutorial de suplemento do Word. Se voc? chegou aqui por meio dos resultados de mecanismos de pesquisa ou por outro link direto, acesse a p?gina de introdu??o do [tutorial de suplemento do Word](../tutorials/word-tutorial.yml) para come?ar pelo in?cio.

Antes de come?ar esta etapa do tutorial, recomendamos a cria??o e manipula??o dos controles de conte?do de Rich Text por meio da interface do usu?rio do Word, para se familiarizar com os controles e suas propriedades. Para saber mais detalhes, confira [Criar formul?rios para preenchimento ou impress?o no Word](https://support.office.com/en-us/article/create-forms-that-users-complete-or-print-in-word-040c5cc1-e309-445b-94ac-542f732c8c8b).

> [!NOTE]
> H? v?rios tipos de controles de conte?do que podem ser adicionados a um documento do Word por meio da interface do usu?rio. Por?m, no momento, s? h? suporte para controles de conte?do de Rich Text no Word.js.


## <a name="create-a-content-control"></a>Criar um controle de conte?do

1. Abra o projeto em seu editor de c?digo. 
2. Abra o arquivo index.html.
3. Abaixo do `div` que cont?m o bot?o `replace-text`, adicione a marca??o a seguir:

    ```html
    <div class="padding">            
        <button class="ms-Button" id="create-content-control">Create Content Control</button>            
    </div>
    ```

4. Abra o arquivo app.js.

5. Abaixo da linha que atribui um identificador de clique ao bot?o `insert-table`, adicione o seguinte c?digo:

    ```js
    $('#create-content-control').click(createContentControl);
    ```

6. Abaixo da fun??o `insertTable`, adicione a fun??o a seguir:

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

7.  Substitua `TODO1` pelo c?digo a seguir. Observa??o:
   - o c?digo tem como objetivo dispor a frase "Office 365" em um controle de conte?do. Para simplificar, ele faz uma pressuposi??o de que a cadeia de caracteres est? presente, e que o usu?rio a selecionou.
   - A propriedade `ContentControl.title` especifica o t?tulo vis?vel do controle de conte?do. 
   - A propriedade `ContentControl.tag` especifica uma marca que pode ser usada para obter uma refer?ncia a um controle de conte?do usando o m?todo `ContentControlCollection.getByTag`, que voc? usar? em uma fun??o posterior. 
   - A propriedade `ContentControl.appearance` especifica a apar?ncia do controle. Usar o valor "Tags" significa que o controle ser? encapsulado entre marcas de abertura e fechamento, e a marca de abertura ter? o t?tulo do controle de conte?do. Outros valores poss?veis s?o "BoundingBox" e "None".
   - A propriedade `ContentControl.color` especifica a cor das marcas ou da borda da caixa delimitadora.

    ```js
    const serviceNameRange = context.document.getSelection();
    const serviceNameContentControl = serviceNameRange.insertContentControl();
    serviceNameContentControl.title = "Service Name";
    serviceNameContentControl.tag = "serviceName";
    serviceNameContentControl.appearance = "Tags";
    serviceNameContentControl.color = "blue";
    ``` 

## <a name="replace-the-content-of-the-content-control"></a>Substituir o conte?do do controle de conte?do

1. Abra o arquivo index.html.
3. Abaixo do `div` que cont?m o bot?o `create-content-control`, adicione a marca??o a seguir:
    ```html
    <div class="padding">            
        <button class="ms-Button" id="replace-content-in-control">Rename Service</button>            
    </div>
    ```

4. Abra o arquivo app.js.

5. Abaixo da linha que atribui um identificador de clique ao bot?o `create-content-control`, adicione o seguinte c?digo:

    ```js
    $('#replace-content-in-control').click(replaceContentInControl);
    ```

6. Abaixo da fun??o `createContentControl`, adicione a fun??o a seguir:

    ```js    fun??o replaceContentInControl() {      Word.run(fun??o) (contexto) {
            
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

7. Replace `TODO1` with the following code. 
    > [!NOTE]
    > The `ContentControlCollection.getByTag` method returns a `ContentControlCollection` of all content controls of the specified tag. We use `getFirst` to get a reference to the desired control.

    ```js
    const serviceNameContentControl = context.document.contentControls.getByTag("serviceName").getFirst();
    serviceNameContentControl.insertText("Fabrikam Online Productivity Suite", "Replace");
    ``` 

## <a name="test-the-add-in"></a>Testar o suplemento

1. Se a janela Git bash ou o prompt de sistema habilitado para Node.JS do tutorial anterior ainda estiverem abertos, digite Ctrl+C duas vezes para interromper a execu??o do servidor Web. Caso contr?rio, abra uma janela Git bash ou um prompt de sistema habilitado para Node.JS e navegue at? a pasta **Iniciar** do projeto.
     > [!NOTE]
     > Embora o servidor de sincroniza??o do navegador recarregue o suplemento no painel de tarefas sempre que voc? fizer uma altera??o em algum arquivo, incluindo o arquivo app.js, ele n?o transcompila o JavaScript, portanto, ? necess?rio repetir o comando de compila??o para que as altera??es em app.js as entrem em vigor. Para fazer isso, interrompa o processo do servidor para que o prompt apare?a e voc? possa inserir o comando de compila??o. Ap?s a compila??o, reinicie o servidor. As pr?ximas etapas executam esse processo.
2. Execute o comando `npm run build` para transcompilar seu c?digo-fonte ES6 para uma vers?o anterior do JavaScript com suporte de todos os hosts nos quais os suplementos do Office podem ser executados.
3. Execute o comando `npm start` para iniciar um servidor Web em um localhost.
4. Feche o painel de tarefas para recarreg?-lo e, no menu **In?cio**, selecione **Mostrar Painel de Tarefas** para reabrir o suplemento.
5. No painel de tarefas, escolha **Inserir Par?grafo** para garantir que haja um par?grafo com "Office 365" no in?cio do documento.
6. Selecione a frase "Office 365" no par?grafo que voc? adicionou e escolha o bot?o **Criar Controle de Conte?do**. A frase est? envolvida por marcas chamadas "Nome do Servi?o".
7. Escolha o bot?o **Renomear Servi?o**. O texto do controle de conte?do muda para "Fabrikam Online Productivity Suite".

    ![Tutorial do Word - Criar o controle de conte?do e alterar seu texto](../images/word-tutorial-content-control.png)
