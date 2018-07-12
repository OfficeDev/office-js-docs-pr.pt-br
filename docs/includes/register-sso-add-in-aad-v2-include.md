

1. Acesse [https://apps.dev.microsoft.com/](https://apps.dev.microsoft.com).

1. Entre com as credenciais de administrador em sua locação do Office 365. Por exemplo, MeuNome@contoso.onmicrosoft.com

1. Clique em **Adicionar um aplicativo**.

1. Quando solicitado, insira **$ ADD-IN-NAME $** como o nome do aplicativo e pressione **Criar aplicativo**.

1. Quando a página de configuração do aplicativo abrir, copie a **ID do aplicativo** e salve-a. Você a usará em um procedimento posterior.

    > [!NOTE]
    > Essa ID é o valor "audience" (público) quando outros aplicativos, como o aplicativo host do Office (por exemplo, PowerPoint, Word, Excel), buscam o acesso autorizado ao aplicativo. Também é a "ID do cliente" do aplicativo quando ela, por sua vez, busca o acesso autorizado ao Microsoft Graph.

1. Na seção **Segredos do Aplicativo**, pressione **Gerar Nova Senha**. Uma caixa de diálogo pop-up abrirá e uma nova senha (também chamada de "segredo do aplicativo") será mostrada. *Copie a senha imediatamente e salve-a com a ID do aplicativo.* Você precisará dela em um procedimento posterior. Feche a caixa de diálogo.

1. Na seção **Plataformas**, clique em **Adicionar plataforma**.

1. Na caixa de diálogo que abrir, selecione **API Web**.

1. A **URI da ID do aplicativo** foi gerada do formulário “api: // $ App ID GUID $”. Insira o **$FQDN-WITHOUT-PROTOCOL$** (com uma barra "/" anexada ao final) entre as barras duplas e o GUID. A ID inteira deve ter o formulário `api://$FQDN-WITHOUT-PROTOCOL$/$App ID GUID$`; por exemplo `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7`.

    > [!NOTE]
    > Se você receber um erro informando que o domínio já tem um dono, mas você é o proprietário, siga o procedimento em [Início rápido: adicionar um nome de domínio personalizado ao Active Directory do Azure](https://docs.microsoft.com/en-us/azure/active-directory/add-custom-domain) para registrá-lo e repita este passo.

    > [!NOTE]
    > A parte do domínio do nome do **Escopo** logo abaixo da **URI da ID do aplicativo** mudará automaticamente, com `/access_as_user` anexado ao final; por exemplo, `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`.

1. Na seção **Aplicativos pré-autorizados**, você identifica os aplicativos que deseja autorizar para o aplicativo da Web do seu suplemento. Cada uma das seguintes IDs precisa ser pré-autorizada. Cada vez que você inserir uma, uma nova caixa de texto vazia aparece. (Insira apenas o GUID.)
    * `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)
    * `57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office Online)
    * `bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Office Online)

1. Abra o menu suspenso do **Escopo** ao lado de cada **ID do aplicativo** e marque a caixa para `api://$FQDN-WITHOUT-PROTOCOL$/$App ID GUID$/access_as_user`.

1. Próximo ao topo da seção **Plataformas**, clique em **Adicionar Plataforma** novamente e selecione **Web**.

1. Na nova seção **Web** em **Plataformas**, insira o seguinte como um **URL de redirecionamento**: `https://$FQDN-WITHOUT-PROTOCOL$`.

1. Role para baixo até a seção **Permissões do Microsoft Graph**, na subseção **Permissões Delegadas**. Use o botão **Adicionar** para abrir a caixa de diálogo **Selecionar Permissões**.

1. Na caixa de diálogo, marque as caixas para `profile` e quaisquer outras permissões do AAD e do Microsoft Graph que seu suplemento precise. Eis alguns exemplos:

    * Files.Read.All
    * offline_access
    * openid
    * perfil

    > [!NOTE]
    > A permissão `User.Read` pode já estar listada por padrão. É uma boa prática não solicitar permissões que não sejam necessárias, portanto, recomendamos que desmarque a caixa para essa permissão se o seu suplemento realmente não precisar dela.

1. Na parte inferior da caixa de diálogo, clique em **OK**.

1. Clique em **Salvar** na parte inferior da página de registro.
