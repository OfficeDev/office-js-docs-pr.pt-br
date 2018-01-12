# <a name="package-your-add-in-using-visual-studio-to-prepare-for-publishing"></a>Empacotar seu suplemento usando o Visual Studio para preparar a publicação

Seu pacote de Suplemento do Office contém um [arquivo de manifesto XML](../overview/add-in-manifests.md) que deve ser usado para publicar o suplemento. Você terá que publicar os arquivos do aplicativo Web do seu projeto separadamente. Este artigo descreve como implantar seu projeto Web e empacotar seu suplemento usando o Visual Studio 2015

## <a name="to-deploy-your-web-project-using-visual-studio-2015"></a>Para implantar seu projeto Web usando o Visual Studio 2015

Conclua as etapas a seguir para implantar seu projeto Web usando o Visual Studio 2015.

1. No **Gerenciador de Soluções**, abra o menu de atalho do projeto do suplemento e escolha **Publicar**.
    
    A página **Publicar seu suplemento** é exibida.
    
2. Na lista suspensa **Perfil atual**, selecione um perfil ou escolha **Novo...** para criar um novo perfil.
    
     >**Observação:**  Um perfil de publicação especifica o servidor que você está implantando, as credenciais necessárias para fazer logon no servidor, os bancos de dados para implantar e outras opções de implantação.

    Se você escolher **Novo...**, o assistente **Criar perfil de publicação** será exibido. Use esse assistente para importar um perfil de publicação de um site de hospedagem, como o Microsoft Azure, ou criar um novo perfil e adicionar seu servidor, as credenciais e outras configurações no procedimento seguinte.
    
    Para mais informações sobre como importar perfis de publicação ou criar novos perfis de publicação, confira [Criar um Perfil de Publicação](http://msdn.microsoft.com/pt-BR/library/dd465337.aspx#creating_a_profile).
    
3. Na página **Publicar seu suplemento**, escolha o link **Implantar seu projeto Web**.
    
    A caixa de diálogo **Publicar na Web** é exibida. Para mais informações sobre como usar esse assistente, confira [Como: Implantar um Projeto da Web usando a Publicação On-Click no Visual Studio](http://msdn.microsoft.com/pt-BR/library/dd465337.aspx).
    

## <a name="to-package-your-add-in-using-visual-studio-2015"></a>Para empacotar seu suplemento usando o Visual Studio 2015

Conclua as etapas a seguir para empacotar seu suplemento usando o Visual Studio 2015.

1. Na página **Publicar seu suplemento**, escolha o link **Empacotar o suplemento**.
    
    O assistente **Publicar Suplementos do Office e do SharePoint** é exibido.
    
2. Na lista suspensa **Onde seu site está hospedado?**, escolha ou digite a URL do site que hospedará os arquivos de conteúdo do seu suplemento e escolha **Concluir**.
    
    Você deve especificar um endereço que comece com o prefixo HTTPS para concluir o assistente. Embora seja geralmente recomendável usar um ponto de extremidade HTTPS para o site, isso não é necessário caso você não pretenda publicar o suplemento na Office Store. Se você quiser usar um ponto de extremidade HTTP para o site, abra o arquivo de manifesto XML em um editor de texto após criar o pacote e substitua o prefixo HTTPS do site por um prefixo HTTP. Confira mais informações em [Por que meus aplicativos e suplementos precisam estar protegidos por SSL?](http://msdn.microsoft.com/pt-BR/library/jj591603#bk_q7).
    
     >**Observação:**  os sites do Azure fornecem automaticamente um ponto de extremidade HTTPS.

    O Visual Studio gera os arquivos nos quais você precisa publicar seu suplemento e, em seguida, abre a pasta de saída de publicação. 
    
Se você pretende enviar seu suplemento à Office Store, pode escolher o link **Executar uma verificação de validação** para identificar problemas que possam impedir a aceitação do seu suplemento. Você deve resolver todos os problemas antes de enviar seu suplemento para o repositório.

Agora você pode carregar seu manifesto XML no local apropriado para [publicar seu suplemento](../publish/publish.md). Você pode encontrar o manifesto XML em `OfficeAppManifests` na pasta `app.publish`. Por exemplo:

 `%UserProfile%\Documents\Visual Studio 2015\Projects\MyApp\bin\Debug\app.publish\OfficeAppManifests`


## <a name="additional-resources"></a>Recursos adicionais



- [Publicar seu Suplemento do Office](../publish/publish.md)
    
- [Enviar Suplementos do SharePoint e do Office e aplicativos Web do Office 365 à Office Store](http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx)
    
