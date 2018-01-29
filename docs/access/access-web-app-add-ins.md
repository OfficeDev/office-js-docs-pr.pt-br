# <a name="create-add-ins-for-access-web-apps"></a>Criar suplementos para aplicativos Web do Access

>**Importante:** Não recomendamos mais criar e usar aplicativos Web do Access e bancos de dados no SharePoint. Como alternativa, use o [Microsoft PowerApps](https://powerapps.microsoft.com/) para criar soluções de negócios sem código para dispositivos móveis e Web.

Este artigo mostra como usar o Visual Studio 2015 para desenvolver um Suplemento do Office destinado a aplicativos Web do Access.

>**Observação:** Para obter informações sobre como desenvolver soluções para o Access usando o VBA, confira [Access](https://msdn.microsoft.com/pt-br/library/fp179695.aspx) na MSDN.

## <a name="prerequisites"></a>Pré-requisitos

Para criar um Suplemento do Office destinado aos aplicativos Web do Access, você precisa:

- Visual Studio 2015

- Um site do SharePoint Online (incluído em várias assinaturas do Office 365). Este site deve ter um catálogo de suplemento. Para saber mais, confira [Configurar um catálogo de suplementos no SharePoint](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).


>**Observação:** Os Suplementos do Office funcionarão com aplicativos Web do Access hospedados no SharePoint Online ou no Office 365. O aplicativo da área de trabalho do Access 2013 não dá suporte aos Suplementos do Office. Os Suplementos do Office direcionados a aplicativos Web do Access são com suporte pela versão 1.1 e posterior do Office.js.


## <a name="create-a-project-in-visual-studio"></a>Criar um projeto no Visual Studio

1.  Abra o Visual Studio no menu, selecione **Arquivo**, **Novo**, **Projeto**. A caixa de diálogo **Novo projeto** será aberta.

2. Na caixa de diálogo **Novo Projeto**, no painel esquerdo, navegue até **Instalado**, **Modelos**, **Visual C#**, **Office/SharePoint**, **Suplementos do Office**.

    >**Observação:**  Se você não tiver este modelo instalado, consulte [O mais recente Microsoft Office Developer Tools para Visual Studio 2015](https://blogs.msdn.microsoft.com/visualstudio/2015/11/23/latest-microsoft-office-developer-tools-for-visual-studio-2015/) para obter informações.

3. Na caixa de diálogo **Novo Projeto**, no painel central, escolha **Suplemento do Office**.

4. Na parte inferior da caixa de diálogo, digite um nome para seu Projeto e selecione **OK**. Isso abrirá a caixa de diálogo **Criar Suplemento do Office**.

5. Na caixa de diálogo **Criar Suplemento do Office**, selecione **Conteúdo** e, em seguida, **Próximo**.

6. Na tela seguinte da caixa de diálogo **Criar Suplemento do Office**, selecione **Suplemento Básico** ou **Suplemento de Visualização de Documentos** e verifique se a caixa de seleção do **Access** está selecionada.

7. Quando terminar, selecione **Concluir**. O Visual Studio criará um projeto inicial para você basear seu trabalho.

8. No **Gerenciador de Soluções**, escolha o projeto da Web do projeto (**nome_do_projeto>Web**). No painel de propriedades, encontre a entrada para **SSL URL**. Ela deve ser similar a: `https://localhost:44314/`. Selecione essa URL e copie-a para sua área de transferência. Você precisará dela em breve.

9. Clique com o botão direito do mouse no nome do seu projeto no **Gerenciador de Soluções**. No menu de contexto, selecione **Publicar**. Isso abrirá o assistente **Publicar seu suplemento**.

10. No assistente **Publicar seu suplemento**, selecione a lista suspensa ao lado de **Perfil atual**. Na lista suspensa, selecione **novo**. Isso abrirá a caixa de diálogo **Publicar Suplementos do Office e do SharePoint**.

11. Na caixa de diálogo selecione **Criar novo perfil**, insira um nome reconhecível para o perfil e, em seguida, escolha **Concluir**. A caixa de diálogo **Publicar Suplementos do Office e do SharePoint** fechará e você retornará ao assistente **Publicar seu suplemento**.

12. No assistente, selecione **Empacotar o suplemento**. Isso finalizará o seu suplemento para que ele possa ser publicado em um catálogo de suplementos no SharePoint.

13. Na página seguinte, para **Onde seu site está hospedado?** insira a URL do host do seu site. Ela pode ser o valor da **URL do SSL** que você copiou na etapa 8. Em seguida, escolha **Concluir**.

14. No **Gerenciador de Soluções**, clique com o botão direito do mouse no nó do Manifesto do projeto (diretamente abaixo do nome do projeto) e selecione **Abrir Pasta no Explorador de Arquivos**. Anote o caminho para esse arquivo. Esse valor será necessário posteriormente.

>**Observação:** Não é possível depurar o suplemento sem implantá-lo com um aplicativo Web do Access.

## <a name="review-the-manifest-and-the-homehtml-file"></a>Analisar o manifesto e o arquivo Home.html

1. No seu projeto do Visual Studio, abra o arquivo **Home.html** e localize as linhas que fazem referência à biblioteca de script do office.js.

    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
    ```

    >**Observação:** As marcas de script fazem referência à versão 1.1 (e posteriores) do Office.js. O Access usará elementos da API introduzidos na versão 1.1.

2. Abra o arquivo do manifesto associado ao seu projeto. Esse arquivo receberá o mesmo nome do seu projeto e terá a extensão ".xml".

3.  No arquivo do manifesto, encontre a seção **Hosts** e procure uma entrada **Host**.

    ```xml
    <Hosts> <Host Name="Database" /> </Hosts>
    ```
    
    >**Observação:** É aqui que os aplicativos que podem usar o suplemento são listados. Como você selecionou **Access** na caixa de diálogo **Criar Suplemento do Office**, o **Banco de Dados** é listado. Se você incluiu o Excel, também haverá uma entrada para **Pasta de Trabalho**.

Os Suplementos do Office e do SharePoint são baseados na Web. O código do suplemento deve ser hospedado em um servidor Web. Neste exemplo, o servidor Web é seu computador de desenvolvimento. O servidor deve estar em execução para servir o suplemento para teste, o que significa que o Visual Studio deve estar executando o suplemento no momento em que você exibi-lo e depurá-lo no SharePoint.

Para um usuário encontrar e utilizar o suplemento, ele precisa estar registrado com um Catálogo de Suplementos no SharePoint. As informações de que o Catálogo de Suplementos precisa estão contidas no arquivo do manifesto.

>**Observação:**  Você precisará criar um aplicativo Web do Access para hospedar seu Suplemento do Office.

## <a name="publish-your-add-in-to-a-sharepoint-online-catalog"></a>Publicar seu suplemento em um catálogo do SharePoint Online

1.  Entre no SharePoint Online ou Office 365 e, em seguida, vá para o **SharePoint Admin Center** escolhendo **Administrador** na barra de ferramentas do Office 365 na parte superior da página.

2. Na página **SharePoint Admin Center**, na barra de links à esquerda, escolha **suplementos**. Isso levará você para o modo de exibição de suplementos.

3. No painel central da página, escolha **Catálogo de Suplementos**. Isso levará você para a página **Catálogo**.

4. Na página **Catálogo**, escolha **Distribuir Suplementos do Office**. Isso leva você a uma página de diretório chamada **Suplementos do Office** que lista todos os Suplementos do Office instalados.

5. Na parte superior da página **Suplementos do Office**, selecione **novo suplemento**. Isso mostrará a caixa de diálogo **Adicionar um documento**.

6. Na caixa de diálogo **Adicionar um documento**, selecione **Procurar** e, em seguida, vá até o local do arquivo de manifesto no seu projeto do Visual Studio. Se você copiou o endereço do seu arquivo de manifesto anteriormente, é possível colá-lo nessa caixa de diálogo.

7. Escolha o arquivo de manifesto no seu projeto e selecione **OK**. Agora, o SharePoint adicionará o seu suplemento à biblioteca local do SharePoint.

>**Observação:**  Esse procedimento presume que você tenha criado um site de teste no SharePoint. Caso ainda não tenha criado, você pode fazê-lo na guia **Sites** na parte superior da janela do SharePoint. Você pode usar um aplicativo Web do Access existente, caso tenha um disponível.

## <a name="create-an-access-web-app-to-host-your-add-in"></a>Criar um aplicativo Web do Access para hospedar seu suplemento

1. Vá para seu site de teste. Na barra de links da esquerda, escolha **Conteúdo do Site**. Isso levará você para a página **Conteúdo do Site** do seu site de teste.

2. Na página **Conteúdo do Site**, escolha **adicionar um suplemento**. Isso levará você para a página **Conteúdo do Site - Seus Suplementos**.

3. Na página **Conteúdo do Site - Seus Suplementos**, use a barra de pesquisa na parte superior da página para procurar **Aplicativo do Access**.

4. Agora você deve ver um bloco para **Aplicativo do Access**.

    >**Observação:**  Lembre-se de que este não é o seu Suplemento do Office. Ele é um novo aplicativo Web do Access. Esses aplicativos Web do Access hospedarão seu Suplemento do Office.

5. Escolher esse novo bloco, exibirá a caixa de diálogo **Adicionar um aplicativo do Access**. Digite um nome exclusivo para seu aplicativo do Access e selecione **Criar**. O SharePoint poderá levar algum tempo para criar seu aplicativo. Quando ele for concluído, você verá seu aplicativo do Access listado na página **Conteúdo do Site** com um rótulo **novo** ao lado dele.

6. O aplicativo do Access agora requer que você o abra na versão para área de trabalho do Microsoft Access 2013 e adicione dados a ele para que, então, seja possível abri-lo e visualizá-lo no SharePoint.

## <a name="add-your-add-in-to-an-access-web-apps"></a>Adicionar seus suplementos a aplicativos Web do Access

1. Abrir aplicativos Web do Access

2. Na barra da guia do SharePoint, selecione o ícone de engrenagem no canto superior esquerdo. Um menu será exibido. Escolha o item do menu **Suplementos do Office**. Isso abrirá a caixa de diálogo **Suplementos do Office**.

3. Escolha o modo de exibição **Minha Organização** e aguarde um momento até que o SharePoint preencha a caixa de diálogo com os Suplementos do Office que estão disponíveis para você.

4. Um dos suplementos na caixa de diálogo deve ser o Suplemento do Office que você registrou em um procedimento anterior. Escolha esse suplemento para inseri-lo em seu Aplicativo Web do Access. Lembre-se que o aplicativo deve estar em execução no Visual Studio para ser detectado e exibido na página do aplicativo Web do Access.

## <a name="debug-your-add-in-for-office"></a>Depurar seus suplementos do Office

Para depurar seu suplemento, no Internet Explorer, pressione F12 ou selecione o ícone de engrenagem na barra da guia do navegador (não o ícone de engrenagem na página do SharePoint). Isso exibirá as ferramentas de depuração F12 fornecidas pelo Internet Explorer 11. Se você estiver usando outro navegador, verifique a documentação do seu navegador para determinar como entrar no modo de depuração.

Neste momento, você pode definir pontos de interrupção, consultar o seu código do JavaScript, explorar o DOM e modificar o código para confirmar se suas alterações aparecem Suplementos do Office direcionados aos aplicativos Web do Access. Confira [Usando as ferramentas F12 de desenvolvedor](http://msdn.microsoft.com/library/ie/bg182326%28v=vs.85%29) para saber mais.

## <a name="next-steps"></a>Próximas etapas

Baixe o exemplo [Office 365: associar e manipular dados em um aplicativo Web do Access](https://code.msdn.microsoft.com/officeapps/Office-365-Bind-and-4876274e) para saber mais sobre como implementar um Suplemento do Office que manipule dados em um aplicativo Web do Access.

## <a name="additional-resources"></a>Recursos adicionais

- [Entender a API JavaScript para suplementos](../develop/understanding-the-javascript-api-for-office.md)

- [API JavaScript para Office](http://dev.office.com/reference/add-ins/javascript-api-for-office)
