# <a name="build-your-first-onenote-add-in"></a>Criar seu primeiro suplemento do OneNote

Neste artigo, voc? passar? pelo processo de criar um suplemento do OneNote usando o jQuery e a API JavaScript para Office.

## <a name="prerequisites"></a>Pr?-requisitos

- [Node.js](https://nodejs.org)

- Instale a ?ltima vers?o do [Yeoman](https://github.com/yeoman/yo) e o [gerador do Yeoman para Suplementos do Office](https://github.com/OfficeDev/generator-office) globalmente.

    ```bash
    npm install -g yo generator-office
    ```

## <a name="create-the-add-in-project"></a>Criar o projeto do suplemento

1. Crie uma pasta na sua unidade local e nomeie-a como `my-onenote-addin`. Esse ? o local em que voc? criar? os arquivos para seu suplemento.

2. Navegue at? a nova pasta.

    ```bash
    cd my-onenote-addin
    ```

3. Use o gerador Yeoman para criar um projeto de suplemento do OneNote. Execute o comando a seguir e responda aos prompts da seguinte forma:

    ```bash
    yo office
    ```

    - **Gostaria de criar uma nova subpasta para o seu projeto?:** `No`
    - **Como deseja nomear seu suplemento?:** `OneNote Add-in`
    - **Para qual aplicativo cliente do Office voc? deseja suporte?:** `OneNote`
    - **Gostaria de criar um novo suplemento?:** `Yes`
    - **Gostaria de usar o TypeScript?:** `No`
    - **Escolha a estrutura:** `Jquery`

    O gerador perguntar? se voc? deseja abrir **resource.html**. N?o ? necess?rio abri-lo para este tutorial, mas fique ? vontade em fazer isso se tiver curiosidade. Escolha Sim ou N?o para concluir o assistente e deixar o gerador fazer seu trabalho.

    ![Uma captura de tela dos prompts e respostas do gerador Yeoman](../images/yo-office-onenote-jquery.png)


## <a name="update-the-code"></a>Atualizar o c?digo

1. No editor de c?digo, abra **index.html** na raiz do projeto. Esse arquivo cont?m o HTML que ser? renderizado no painel de tarefas do suplemento.

2. Substitua o elemento `<main>` dentro do elemento `<body>` com a marca??o a seguir e salve o arquivo. Isso adiciona uma ?rea de texto e um bot?o usando [componentes do Office UI Fabric](http://dev.office.com/fabric/components).

    ```html
    <main class="ms-welcome__main">
        <br />
        <p class="ms-font-l">Enter content below</p>
        <div class="ms-TextField ms-TextField--placeholder">
            <textarea id="textBox" rows="5"></textarea>
        </div>
        <button id="addOutline" class="ms-welcome__action ms-Button ms-Button--hero ms-u-slideUpIn20">
            <span class="ms-Button-label">Add Outline</span>
            <span class="ms-Button-icon"><i class="ms-Icon"></i></span>
            <span class="ms-Button-description">Adds the content above to the current page.</span>
        </button>
    </main>
    ```

3. Abra o arquivo **app.js** para especificar o script do suplemento. Substitua todo o conte?do pelo c?digo a seguir e salve o arquivo.

    ```js
    'use strict';

    (function () {

        Office.initialize = function (reason) {
            $(document).ready(function () {
                app.initialize();

                // Set up event handler for the UI.
                $('#addOutline').click(addOutlineToPage);
            });
        };

        // Add the contents of the text area to the page.
        function addOutlineToPage() {        
            OneNote.run(function (context) {
                var html = '<p>' + $('#textBox').val() + '</p>';

                // Get the current page.
                var page = context.application.getActivePage();

                // Queue a command to load the page with the title property.             
                page.load('title'); 

                // Add an outline with the specified HTML to the page.
                var outline = page.addOutline(40, 90, html);

                // Run the queued commands, and return a promise to indicate task completion.
                return context.sync()
                    .then(function() {
                        console.log('Added outline to page ' + page.title);
                    })
                    .catch(function(error) {
                        app.showNotification("Error: " + error); 
                        console.log("Error: " + error); 
                        if (error instanceof OfficeExtension.Error) { 
                            console.log("Debug info: " + JSON.stringify(error.debugInfo)); 
                        } 
                    }); 
            });
        }
    })();
    ```

## <a name="update-the-manifest"></a>Atualizar o manifesto

1. Abra o arquivo **one-note-add-in-manifest.xml** para definir as configura??es e os recursos do suplemento.

2. O elemento `ProviderName` tem um valor de espa?o reservado. Substitua-o com seu nome.

3. O atributo `DefaultValue` do elemento `Description` tem um espa?o reservado. Substitua-o por **um suplemento do painel de tarefas do OneNote**.

4. Salve o arquivo.

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="OneNote Add-in" />
    <Description DefaultValue="A task pane add-in for OneNote"/>
    ...
    ```

## <a name="start-the-dev-server"></a>Iniciar o servidor de desenvolvimento

[!include[Start server section](../includes/quickstart-yo-start-server.md)]

## <a name="try-it-out"></a>Experimente

1. No [OneNote Online](https://www.onenote.com/notebooks), abra um bloco de anota??es.

2. Escolha **Inserir > Suplementos do Office** para abrir a caixa de di?logo Suplementos do Office.

    - Se voc? estiver conectado ? sua conta de consumidor, selecione a guia **MEUS SUPLEMENTOS** e escolha  **Carregar Meu Suplemento**.

    - Se voc? estiver conectado ? sua conta corporativa ou de estudante, selecione a guia **MINHA ORGANIZA??O** e escolha  **Carregar Meu Suplemento**. 

    A imagem a seguir mostra a guia **MEUS SUPLEMENTOS** para blocos de anota??es do consumidor.

    <img alt="The Office Add-ins dialog showing the MY ADD-INS tab" src="../images/onenote-office-add-ins-dialog.png" width="500">

3. Na caixa de di?logo Carregar suplemento, navegue at? **one-note-add-in-manifest.xml** na pasta do projeto e escolha **Carregar**. 

4. Na guia **P?gina Inicial**, escolha o bot?o **Exibir painel de tarefas** na faixa de op??es. O painel de tarefas do suplemento abre em um iFrame perto da p?gina do OneNote.

5. Insira algum texto na ?rea de texto e escolha **Adicionar estrutura de t?picos**. O texto inserido ? adicionado ? pagina. 

    ![O suplemento do OneNote criado a partir deste passo a passo](../images/onenote-first-add-in.png)

## <a name="troubleshooting-and-tips"></a>Dicas e solu??o de problemas

- Voc? pode depurar o suplemento usando as ferramentas de desenvolvedor do seu navegador. Quando voc? estiver usando o servidor Web Gulp e depurando no Internet Explore ou no Chrome, voc? pode salvar as altera??es localmente e apenas atualize o iFrame do suplemento.

- Quando voc? inspecionar um objeto do OneNote, as propriedades que est?o atualmente dispon?veis usam valores reais de exibi??o. As propriedades que precisam ser carregadas exibem *undefined*. Expanda o n? `_proto_` para ver as propriedades definidas no objeto, mas que ainda n?o foram carregadas.

   ![Carregar o objeto do OneNote em um depurador](../images/onenote-debug.png)

- Voc? precisa habilitar conte?do misto no navegador, se o seu suplemento usar todos os recursos HTTP. Os suplementos de produ??o devem usar apenas recursos HTTPS seguros.

- ? poss?vel abrir os suplementos do Painel de Tarefas em praticamente qualquer lugar, mas os suplementos de conte?do podem ser inseridos apenas no conte?do normal da p?gina (ou seja, fora t?tulos, imagens, iFrames, etc.). 

## <a name="next-steps"></a>Pr?ximas etapas

Parab?ns, voc? criou com ?xito um suplemento do OneNote! Em seguida, saiba mais sobre os principais conceitos de cria??o de suplementos do OneNote.

> [!div class="nextstepaction"]
> [Vis?o geral da programa??o da API JavaScript do OneNote](../onenote/onenote-add-ins-programming-overview.md)

## <a name="see-also"></a>Veja tamb?m

- [Vis?o geral da programa??o da API JavaScript do OneNote](../onenote/onenote-add-ins-programming-overview.md)
- [Refer?ncia da API JavaScript do OneNote](https://dev.office.com/reference/add-ins/onenote/onenote-add-ins-javascript-reference)
- [Amostra de Rubric Grader](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Vis?o geral da plataforma Suplementos do Office](../overview/office-add-ins.md)
