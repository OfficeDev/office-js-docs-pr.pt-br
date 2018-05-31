Nesta etapa do tutorial, você aprenderá a criar controles de conteúdo de Rich Text no documento e, depois, como inserir e substituir conteúdo nos controles. 

> [!NOTE]
> Esta página descreve uma etapa individual de um tutorial de suplemento do Word. Se você chegou aqui por meio dos resultados de mecanismos de pesquisa ou por outro link direto, acesse a página de introdução do [tutorial de suplemento do Word](../tutorials/word-tutorial.yml) para começar pelo início.

Antes de começar esta etapa do tutorial, recomendamos a criação e manipulação dos controles de conteúdo de Rich Text por meio da interface do usuário do Word, para se familiarizar com os controles e suas propriedades. Para saber mais detalhes, confira [Criar formulários para preenchimento ou impressão no Word](https://support.office.com/en-us/article/create-forms-that-users-complete-or-print-in-word-040c5cc1-e309-445b-94ac-542f732c8c8b).

> [!NOTE]
> Há vários tipos de controles de conteúdo que podem ser adicionados a um documento do Word por meio da interface do usuário. Porém, no momento, só há suporte para controles de conteúdo de Rich Text no Word.js.


## <a name="create-a-content-control"></a>Criar um controle de conteúdo

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

7.  Substitua `TODO1` pelo código a seguir. Observação:
   - o código tem como objetivo dispor a frase "Office 365" em um controle de conteúdo. Para simplificar, ele faz uma pressuposição de que a cadeia de caracteres está presente, e que o usuário a selecionou.
   - A propriedade `ContentControl.title` especifica o título visível do controle de conteúdo. 
   - A propriedade `ContentControl.tag` especifica uma marca que pode ser usada para obter uma referência a um controle de conteúdo usando o método `ContentControlCollection.getByTag`, que você usará em uma função posterior. 
   - A propriedade `ContentControl.appearance` especifica a aparência do controle. Usar o valor "Tags" significa que o controle será encapsulado entre marcas de abertura e fechamento, e a marca de abertura terá o título do controle de conteúdo. Outros valores possíveis são "BoundingBox" e "None".
   - A propriedade `ContentControl.color` especifica a cor das marcas ou da borda da caixa delimitadora.

    ```js
    const serviceNameRange = context.document.getSelection();
    const serviceNameContentControl = serviceNameRange.insertContentControl();
    serviceNameContentControl.title = "Service Name";
    serviceNameContentControl.tag = "serviceName";
    serviceNameContentControl.appearance = "Tags";
    serviceNameContentControl.color = "blue";
    ``` 

## <a name="replace-the-content-of-the-content-control"></a>Substituir o conteúdo do controle de conteúdo

1. Abra o arquivo index.html.
3. Abaixo do `div` que contém o botão `create-content-control`, adicione a marcação a seguir:
    ```html
    <div class="padding">            
        <button class="ms-Button" id="replace-content-in-control">Rename Service</button>            
    </div>
    ```

4. Abra o arquivo app.js.

5. Abaixo da linha que atribui um identificador de clique ao botão `create-content-control`, adicione o seguinte código:

    ```js
    $('#replace-content-in-control').click(replaceContentInControl);
    ```

6. Abaixo da função `createContentControl`, adicione a função a seguir:

    ```js    função replaceContentInControl() {      Word.run(função) (contexto) {
            
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

1. Se a janela Git bash ou o prompt de sistema habilitado para Node.JS do tutorial anterior ainda estiverem abertos, digite Ctrl+C duas vezes para interromper a execução do servidor Web. Caso contrário, abra uma janela Git bash ou um prompt de sistema habilitado para Node.JS e navegue até a pasta **Iniciar** do projeto.
     > [!NOTE]
     > Embora o servidor de sincronização do navegador recarregue o suplemento no painel de tarefas sempre que você fizer uma alteração em algum arquivo, incluindo o arquivo app.js, ele não transcompila o JavaScript, portanto, é necessário repetir o comando de compilação para que as alterações em app.js as entrem em vigor. Para fazer isso, interrompa o processo do servidor para que o prompt apareça e você possa inserir o comando de compilação. Após a compilação, reinicie o servidor. As próximas etapas executam esse processo.
2. Execute o comando `npm run build` para transcompilar seu código-fonte ES6 para uma versão anterior do JavaScript com suporte de todos os hosts nos quais os suplementos do Office podem ser executados.
3. Execute o comando `npm start` para iniciar um servidor Web em um localhost.
4. Feche o painel de tarefas para recarregá-lo e, no menu **Início**, selecione **Mostrar Painel de Tarefas** para reabrir o suplemento.
5. No painel de tarefas, escolha **Inserir Parágrafo** para garantir que haja um parágrafo com "Office 365" no início do documento.
6. Selecione a frase "Office 365" no parágrafo que você adicionou e escolha o botão **Criar Controle de Conteúdo**. A frase está envolvida por marcas chamadas "Nome do Serviço".
7. Escolha o botão **Renomear Serviço**. O texto do controle de conteúdo muda para "Fabrikam Online Productivity Suite".

    ![Tutorial do Word - Criar o controle de conteúdo e alterar seu texto](../images/word-tutorial-content-control.png)
