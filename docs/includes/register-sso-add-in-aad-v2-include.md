

1. Navegar para [https://apps.dev.microsoft.com/](https://apps.dev.microsoft.com).

1. Entre com as credenciais de administrador em seu locat?rio do Office 365. Por exemplo: MeuNome@contoso.onmicrosoft.com

1. Clique em **Adicionar um aplicativo**.

1. Quando solicitado, digite **$ADD-IN-NAME$** como o nome do aplicativo e pressione **Criar aplicativo**.

1. Quando a p?gina de configura??o do aplicativo abrir, copie a **ID do aplicativo** e salve-a. Voc? a usar? em um procedimento posterior.

    > [!NOTE]
    > Essa ID ? o valor "audience" (p?blico) quando outros aplicativos, como o aplicativo host do Office (por exemplo, PowerPoint, Word, Excel), buscam o acesso autorizado ao aplicativo. Tamb?m ? a "ID do cliente" do aplicativo quando ela, por sua vez, busca o acesso autorizado ao Microsoft Graph.

1. Na se??o **Segredos do Aplicativo**, pressione **Gerar Nova Senha**. Uma caixa de di?logo pop-up abrir? e uma nova senha (tamb?m chamada de "segredo do aplicativo") ser? mostrada. *Copie a senha imediatamente e salve-a com a ID do aplicativo.* Voc? precisar? dela em um procedimento posterior. Feche a caixa de di?logo.

1. Na se??o **Plataformas**, clique em **Adicionar plataforma**.

1. Na caixa de di?logo que abrir, selecione **API Web**.

1. A **URI da ID do aplicativo** foi gerada do formul?rio ?api: // $ App ID GUID $?. Insira o **$FQDN-WITHOUT-PROTOCOL$** (com uma barra "/" anexada ao final) entre as barras duplas e o GUID. A ID inteira deve ter o formul?rio `api://$FQDN-WITHOUT-PROTOCOL$/$App ID GUID$`; por exemplo `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7`.

    > [!NOTE]
    > Se voc? receber um erro informando que o dom?nio j? tem um dono, mas voc? ? o propriet?rio, siga o procedimento em [In?cio r?pido: adicionar um nome de dom?nio personalizado ao Active Directory do Azure](https://docs.microsoft.com/en-us/azure/active-directory/add-custom-domain) para registr?-lo e repita este passo.

    > [!NOTE]
    > A parte do dom?nio do nome do **Escopo** logo abaixo da **URI da ID do aplicativo** mudar? automaticamente para corresponder, com `/access_as_user` anexado ao final; por exemplo, `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`.

1. Na se??o **Aplicativos pr?-autorizados** , voc? identifica os aplicativos que deseja autorizar para o aplicativo da Web do seu suplemento. Cada uma das seguintes IDs precisa ser pr?-autorizada. Cada vez que voc? inserir uma, uma nova caixa de texto vazia aparece. (Insira apenas o GUID.)
    * `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)
    * `57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office Online)
    * `bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Office Online)

1. Abra o menu suspenso do **Escopo** ao lado de cada **ID do aplicativo** e marque a caixa para `api://$FQDN-WITHOUT-PROTOCOL$/$App ID GUID$/access_as_user`.

1. Pr?ximo ao topo da se??o **Plataformas**, clique em **Adicionar Plataforma** novamente e selecione **Web**.

1. Na nova se??o **Web** em **Plataformas**, insira o seguinte como um **URL de redirecionamento**: `https://$FQDN-WITHOUT-PROTOCOL$`.

1. Role para baixo at? a se??o **Permiss?es do Microsoft Graph**, na subse??o **Permiss?es Delegadas**. Use o bot?o **Adicionar** para abrir a caixa de di?logo **Selecionar Permiss?es**.

1. Na caixa de di?logo, marque as caixas para `profile` e quaisquer outras permiss?es do AAD e do Microsoft Graph que seu suplemento precise. Eis alguns exemplos:

    * Files.Read.All
    * offline_access
    * openid
    * perfil

    > [!NOTE]
    > A permiss?o `User.Read` pode j? estar listada por padr?o. ? uma boa pr?tica n?o solicitar permiss?es que n?o sejam necess?rias, portanto, recomendamos que desmarque a caixa para essa permiss?o se o seu suplemento realmente n?o precisar.

1. Na parte inferior da caixa de di?logo, clique em **OK**.

1. Clique em**Salvar** na parte inferior da p?gina de registro.
