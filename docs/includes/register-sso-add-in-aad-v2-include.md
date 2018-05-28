

1. <span data-ttu-id="f44b0-101">Navegar para [https://apps.dev.microsoft.com/](https://apps.dev.microsoft.com).</span><span class="sxs-lookup"><span data-stu-id="f44b0-101">Navigate to [site\wwwroothttps://apps.dev.microsoft.com/[nameofyourazurefunction]](https://apps.dev.microsoft.com)</span></span>

1. <span data-ttu-id="f44b0-p101">Entre com as credenciais de administrador em seu locat?rio do Office 365. Por exemplo: MeuNome@contoso.onmicrosoft.com</span><span class="sxs-lookup"><span data-stu-id="f44b0-p101">Sign-in with the admin credentials to your Office 365 tenancy. For example, MyName@contoso.onmicrosoft.com</span></span>

1. <span data-ttu-id="f44b0-104">Clique em **Adicionar um aplicativo**.</span><span class="sxs-lookup"><span data-stu-id="f44b0-104">Click **Add an app**.</span></span>

1. <span data-ttu-id="f44b0-105">Quando solicitado, digite **$ADD-IN-NAME$** como o nome do aplicativo e pressione **Criar aplicativo**.</span><span class="sxs-lookup"><span data-stu-id="f44b0-105">When prompted, use ?Office-Add-in-ASPNET-SSO? as the app name, and then press Create application.</span></span>

1. <span data-ttu-id="f44b0-p102">Quando a p?gina de configura??o do aplicativo abrir, copie a **ID do aplicativo** e salve-a. Voc? a usar? em um procedimento posterior.</span><span class="sxs-lookup"><span data-stu-id="f44b0-p102">When the configuration page for the app opens, copy the **Application Id** and save it. You'll use it in a later procedure.</span></span>

    > [!NOTE]
    > <span data-ttu-id="f44b0-p103">Essa ID ? o valor "audience" (p?blico) quando outros aplicativos, como o aplicativo host do Office (por exemplo, PowerPoint, Word, Excel), buscam o acesso autorizado ao aplicativo. Tamb?m ? a "ID do cliente" do aplicativo quando ela, por sua vez, busca o acesso autorizado ao Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="f44b0-p103">This ID is the ?audience? value when other applications, such as the Office host application (e.g., PowerPoint, Word, Excel), seek authorized access to the application. It is also the ?client ID? of the application when it, in turn, seeks authorized access to Microsoft Graph.</span></span>

1. <span data-ttu-id="f44b0-p104">Na se??o **Segredos do Aplicativo**, pressione **Gerar Nova Senha**. Uma caixa de di?logo pop-up abrir? e uma nova senha (tamb?m chamada de "segredo do aplicativo") ser? mostrada. *Copie a senha imediatamente e salve-a com a ID do aplicativo.* Voc? precisar? dela em um procedimento posterior. Feche a caixa de di?logo.</span><span class="sxs-lookup"><span data-stu-id="f44b0-p104">In the **Application Secrets** section, press **Generate New Password**. A popup dialog opens with a new password (also called an ?app secret?) displayed. *Copy the password immediately and save it with the application ID.* You'll need it in a later procedure. Then close the dialog.</span></span>

1. <span data-ttu-id="f44b0-115">Na se??o **Plataformas**, clique em **Adicionar plataforma**.</span><span class="sxs-lookup"><span data-stu-id="f44b0-115">In the **Platforms** section, click **Add Platform**.</span></span>

1. <span data-ttu-id="f44b0-116">Na caixa de di?logo que abrir, selecione **API Web**.</span><span class="sxs-lookup"><span data-stu-id="f44b0-116">In the dialog that opens, select **Web API**.</span></span>

1. <span data-ttu-id="f44b0-117">A **URI da ID do aplicativo** foi gerada do formul?rio ?api: // $ App ID GUID $?.</span><span class="sxs-lookup"><span data-stu-id="f44b0-117">An **Application ID URI** has been generated of the form ?api://$App ID GUID$?.</span></span> <span data-ttu-id="f44b0-118">Insira o **$FQDN-WITHOUT-PROTOCOL$** (com uma barra "/" anexada ao final) entre as barras duplas e o GUID.</span><span class="sxs-lookup"><span data-stu-id="f44b0-118">Insert the **$FQDN-WITHOUT-PROTOCOL$** (with a forward slash "/" appended to the end) between the double forward slashes and the GUID.</span></span> <span data-ttu-id="f44b0-119">A ID inteira deve ter o formul?rio `api://$FQDN-WITHOUT-PROTOCOL$/$App ID GUID$`; por exemplo `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7`.</span><span class="sxs-lookup"><span data-stu-id="f44b0-119">The entire ID should have the form `api://$FQDN-WITHOUT-PROTOCOL$/$App ID GUID$`; for example `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7`.</span></span>

    > [!NOTE]
    > <span data-ttu-id="f44b0-120">Se voc? receber um erro informando que o dom?nio j? tem um dono, mas voc? ? o propriet?rio, siga o procedimento em [In?cio r?pido: adicionar um nome de dom?nio personalizado ao Active Directory do Azure](https://docs.microsoft.com/en-us/azure/active-directory/add-custom-domain) para registr?-lo e repita este passo.</span><span class="sxs-lookup"><span data-stu-id="f44b0-120">If you get an error saying that the domain is already owned, but you own it, follow the procedure at [Quickstart: Add a custom domain name to Azure Active Directory](https://docs.microsoft.com/en-us/azure/active-directory/add-custom-domain) to register it, and then repeat this step.</span></span>

    > [!NOTE]
    > <span data-ttu-id="f44b0-121">A parte do dom?nio do nome do **Escopo** logo abaixo da **URI da ID do aplicativo** mudar? automaticamente para corresponder, com `/access_as_user` anexado ao final; por exemplo, `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`.</span><span class="sxs-lookup"><span data-stu-id="f44b0-121">The domain part of the **Scope** name just below the **Application ID URI** will automatically change to match.</span></span>

1. <span data-ttu-id="f44b0-122">Na se??o **Aplicativos pr?-autorizados** , voc? identifica os aplicativos que deseja autorizar para o aplicativo da Web do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="f44b0-122">In the **Pre-authorized applications** section, you identify the applications that you want to authorize to your add-in's web application.</span></span> <span data-ttu-id="f44b0-123">Cada uma das seguintes IDs precisa ser pr?-autorizada.</span><span class="sxs-lookup"><span data-stu-id="f44b0-123">Each of the following IDs needs to be pre-authorized.</span></span> <span data-ttu-id="f44b0-124">Cada vez que voc? inserir uma, uma nova caixa de texto vazia aparece.</span><span class="sxs-lookup"><span data-stu-id="f44b0-124">Each time you enter one, a new empty textbox appears.</span></span> <span data-ttu-id="f44b0-125">(Insira apenas o GUID.)</span><span class="sxs-lookup"><span data-stu-id="f44b0-125">(Enter only the GUID.)</span></span>
    * <span data-ttu-id="f44b0-126">`d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)</span><span class="sxs-lookup"><span data-stu-id="f44b0-126">`d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)</span></span>
    * <span data-ttu-id="f44b0-127">`57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office Online)</span><span class="sxs-lookup"><span data-stu-id="f44b0-127">`57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office Online)</span></span>
    * <span data-ttu-id="f44b0-128">`bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Office Online)</span><span class="sxs-lookup"><span data-stu-id="f44b0-128">`bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Office Online)</span></span>

1. <span data-ttu-id="f44b0-129">Abra o menu suspenso do **Escopo** ao lado de cada **ID do aplicativo** e marque a caixa para `api://$FQDN-WITHOUT-PROTOCOL$/$App ID GUID$/access_as_user`.</span><span class="sxs-lookup"><span data-stu-id="f44b0-129">Open the **Scope** drop-down beside each **Application ID** and check the box for `api://$FQDN-WITHOUT-PROTOCOL$/$App ID GUID$/access_as_user`.</span></span>

1. <span data-ttu-id="f44b0-130">Pr?ximo ao topo da se??o **Plataformas**, clique em **Adicionar Plataforma** novamente e selecione **Web**.</span><span class="sxs-lookup"><span data-stu-id="f44b0-130">Near the top of the **Platforms** section, click **Add Platform** again and select **Web**.</span></span>

1. <span data-ttu-id="f44b0-131">Na nova se??o **Web** em **Plataformas**, insira o seguinte como um **URL de redirecionamento**: `https://$FQDN-WITHOUT-PROTOCOL$`.</span><span class="sxs-lookup"><span data-stu-id="f44b0-131">In the new **Web** section under **Platforms**, enter the following as a **Redirect URL**: `https://$FQDN-WITHOUT-PROTOCOL$`.</span></span>

1. <span data-ttu-id="f44b0-p107">Role para baixo at? a se??o **Permiss?es do Microsoft Graph**, na subse??o **Permiss?es Delegadas**. Use o bot?o **Adicionar** para abrir a caixa de di?logo **Selecionar Permiss?es**.</span><span class="sxs-lookup"><span data-stu-id="f44b0-p107">Scroll down to the **Microsoft Graph Permissions** section, the **Delegated Permissions** subsection. Use the **Add** button to open a **Select Permissions** dialog.</span></span>

1. <span data-ttu-id="f44b0-134">Na caixa de di?logo, marque as caixas para `profile` e quaisquer outras permiss?es do AAD e do Microsoft Graph que seu suplemento precise.</span><span class="sxs-lookup"><span data-stu-id="f44b0-134">In the dialog box, check the boxes for `profile` and any other AAD and Microsoft Graph permissions that your add-in needs.</span></span> <span data-ttu-id="f44b0-135">Eis alguns exemplos:</span><span class="sxs-lookup"><span data-stu-id="f44b0-135">The following are examples:</span></span>

    * <span data-ttu-id="f44b0-136">Files.Read.All</span><span class="sxs-lookup"><span data-stu-id="f44b0-136">Files.Read.All</span></span>
    * <span data-ttu-id="f44b0-137">offline_access</span><span class="sxs-lookup"><span data-stu-id="f44b0-137">offline_access</span></span>
    * <span data-ttu-id="f44b0-138">openid</span><span class="sxs-lookup"><span data-stu-id="f44b0-138">openid</span></span>
    * <span data-ttu-id="f44b0-139">perfil</span><span class="sxs-lookup"><span data-stu-id="f44b0-139">profile</span></span>

    > [!NOTE]
    > <span data-ttu-id="f44b0-140">A permiss?o `User.Read` pode j? estar listada por padr?o.</span><span class="sxs-lookup"><span data-stu-id="f44b0-140">The `User.Read` permission may already be listed by default.</span></span> <span data-ttu-id="f44b0-141">? uma boa pr?tica n?o solicitar permiss?es que n?o sejam necess?rias, portanto, recomendamos que desmarque a caixa para essa permiss?o se o seu suplemento realmente n?o precisar.</span><span class="sxs-lookup"><span data-stu-id="f44b0-141">It is a good practice not to ask for permissions that are not needed, so we recommend that you uncheck the box for this permission.</span></span>

1. <span data-ttu-id="f44b0-142">Na parte inferior da caixa de di?logo, clique em **OK**.</span><span class="sxs-lookup"><span data-stu-id="f44b0-142">At the bottom of the dialog, click **OK**.</span></span>

1. <span data-ttu-id="f44b0-143">Clique em**Salvar** na parte inferior da p?gina de registro.</span><span class="sxs-lookup"><span data-stu-id="f44b0-143">At the bottom of the registration page, click **Save**.</span></span>
