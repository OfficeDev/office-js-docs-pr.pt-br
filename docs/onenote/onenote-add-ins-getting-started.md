# Crie seu primeiro suplemento do OneNote
<a id="build-your-first-onenote-add-in" class="xliff"></a>

Este artigo ajuda você a criar um suplemento simples de painel de tarefas que adiciona texto a uma página do OneNote.

A imagem a seguir mostra o suplemento que você criará.

   ![O suplemento do OneNote criado a partir deste passo a passo](../../images/onenote-first-add-in.png)

<a name="setup"></a>
## Etapa 1: Configurar o seu ambiente de desenvolvimento e criar um projeto de suplemento
<a id="step-1-set-up-your-dev-environment-and-create-an-add-in-project" class="xliff"></a>
Siga as instruções do artigo [Criar um suplemento do Office usando um editor](../get-started/create-an-office-add-in-using-any-editor.md) para instalar os pré-requisitos necessários e executar o gerador Yeoman do Office a fim de criar um novo projeto de suplemento. A tabela a seguir enumera os atributos de projeto que devem ser selecionados no gerador Yeoman.

| Opção | Valor |
|:------|:------|
| Novas subpastas | (aceitar o padrão) |
| Nome do suplemento | Suplemento do OneNote |
| Aplicativo Office compatível | (selecionar OneNote) |
| Criar novo suplemento | Sim, quero um novo suplemento |
| Adicionar [TypeScript](https://www.typescriptlang.org/) | Não |
| Escolher estrutura | Jquery |

<a name="develop"></a>
## Etapa 2: modificar o suplemento
<a id="step-2-modify-the-add-in" class="xliff"></a>
Você pode editar o suplemento usando um editor de texto ou IDE. Se ainda não experimentou o Visual Studio Code, você pode [baixá-lo gratuitamente](https://code.visualstudio.com/) no Windows, no Mac OSX e no Linux.

1 – Abra **index.html** no diretório do projeto. 

2 – Substitua o elemento `<main>` pelo seguinte código. Isso adiciona uma área de texto e um botão usando [componentes do Office UI Fabric](http://dev.office.com/fabric/components).

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

3 – Abra **app.js** (ou app.ts, se estiver usando o TypeScript) no diretório do projeto. Edite a função **Office.initialize** para adicionar um evento de clique ao botão **Adicionar estrutura de tópicos**, da seguinte maneira.

```js
// The initialize function is run each time the page is loaded.
Office.initialize = function (reason) {
   $(document).ready(function () {
       app.initialize();
       
       // Set up event handler for the UI.
       $('#addOutline').click(addOutlineToPage);
   });
};
```
 
4 – Substitua o método **run** pelo seguinte método **addOutlineToPage**. Isto adiciona o conteúdo de uma área de texto à página.

```js
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
```

<a name="test"></a>
## Etapa 3: testar o suplemento no OneNote Online
<a id="step-3-test-the-add-in-on-onenote-online" class="xliff"></a>
1 – Inicie o servidor HTTPS.  

  a. Abra um prompt **cmd** / Terminal e vá para a pasta do projeto de suplemento. 
  
  b. Execute o comando, conforme mostrado abaixo.

  ```
  C:\your-local-path\onenote add-in\> npm start
  ```

2 – Instale o certificado autoassinado como um certificado confiável. Você só precisa fazer isso uma vez no seu computador para todos os projetos de suplemento criados com o gerador Yeoman do Office. Saiba mais em [Adicionar certificados autoassinados como certificado raiz de confiança](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).

3 – Vá para o [OneNote Online](https://www.onenote.com/notebooks) e abra um bloco de anotações.

4 – Escolha **Inserir > Suplementos do Office**. Isso abre a caixa de diálogo Suplementos do Office.

  -Se você estiver conectado com a sua conta de consumidor, escolha a guia **MEUS SUPLEMENTOS** e, em seguida, escolha  **Carregar Meu Suplemento**.
  
  -Se você estiver conectado com a sua conta corporativa ou de estudante, escolha a guia **MINHA ORGANIZAÇÃO** e, em seguida, escolha **Carregar Meu Suplemento**. 
  
  A imagem a seguir mostra a guia **MEUS SUPLEMENTOS** para blocos de anotações do consumidor.

  ![O diálogo Suplementos do Office mostrando a guia MEUS SUPLEMENTOS](../../images/onenote-office-add-ins-dialog.png)

5 – No diálogo Carregar suplemento, navegue até **onenote-add-in-manifest.xml** na pasta do projeto e escolha **Carregar**. O arquivo do manifesto será colocado no armazenamento local do navegador durante o teste.

6 – O suplemento abre em um iFrame ao lado da página do OneNote. Insira algum texto na área correspondente e escolha **Adicionar estrutura de tópicos**. O texto inserido é adiciona à pagina. 

## Dicas e solução de problemas
<a id="troubleshooting-and-tips" class="xliff"></a>
-Você pode depurar o suplemento usando as ferramentas de desenvolvedor do seu navegador. Quando você estiver usando o servidor Web Gulp e depurando no Internet Explore ou no Chrome, você pode salvar as alterações localmente e apenas atualize o iFrame do suplemento.

-Quando você inspecionar um objeto do OneNote, as propriedades que estão atualmente disponíveis usam valores reais de exibição. As propriedades que precisam ser carregadas exibem *indefinido*. Expanda o nó `_proto_` para ver as propriedades definidas no objeto, mas que ainda não foram carregadas.

![Carregar o objeto do OneNote em um depurador](../../images/onenote-debug.png)

Você precisa habilitar conteúdo misto no navegador, se o seu suplemento usar todos os recursos HTTP. Os suplementos de produção devem usar apenas recursos HTTPS seguros.

É possível abrir os suplementos do Painel de Tarefas em praticamente qualquer lugar, mas os suplementos de conteúdo podem ser inseridos apenas no conteúdo normal da página (ou seja, fora títulos, imagens, iFrames, etc.). 

## Recursos Adicionais
<a id="additional-resources" class="xliff"></a>

- [Visão geral sobre a programação da API JavaScript do OneNote](onenote-add-ins-programming-overview.md)
- [Referência da API JavaScript do OneNote](../../reference/onenote/onenote-add-ins-javascript-reference.md)
- [Amostra de Rubric Grader](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Visão geral da plataforma de Suplementos do Office](https://dev.office.com/docs/add-ins/overview/office-add-ins)
