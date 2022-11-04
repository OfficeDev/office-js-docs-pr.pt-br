## <a name="register-the-add-in-with-microsoft-identity-platform"></a>Registrar o suplemento com plataforma de identidade da Microsoft

Você precisa criar um registro de aplicativo no Azure que represente seu servidor Web. Isso permite o suporte à autenticação para que tokens de acesso adequados possam ser emitidos para o código do cliente no JavaScript. Esse registro dá suporte ao SSO no cliente e à autenticação de fallback usando a MSAL (Biblioteca de Autenticação da Microsoft).


1. Entre no [portal do Azure](https://portal.azure.com/) com as credenciais ***admin** _ para sua locação do Microsoft 365. Por exemplo, _*MyName@contoso.onmicrosoft.com**.
1. Selecione **Registros de aplicativos**. Se você não vir o ícone, pesquise por "registro de aplicativo" na barra de pesquisa.

    :::image type="content" source="../images/azure-portal-select-app-registration.png" alt-text="A página inicial portal do Azure.":::

    A página **Registros de aplicativo** é exibida.

1. Selecione **Novo registro**.

    :::image type="content" source="../images/azure-portal-select-new-registration.png" alt-text="Novo registro no painel Registros de aplicativo.":::

    O **painel Registrar um aplicativo** é exibido.

1. Em **Gerenciar**, selecione **Registros de aplicativo** >  **Novo registro**. No painel **Registrar um aplicativo** , defina os valores da seguinte maneira.

    * Defina **Nome** para `<add-in-name>`.
    * Defina **tipos de conta com suporte** **como Contas em qualquer diretório organizacional (qualquer diretório Azure AD - multilocatário) e contas pessoais da Microsoft (por exemplo, Skype, Xbox)**.
    * Defina **o URI de redirecionamento** para usar a plataforma `<redirect-platform>` e o URI como `<redirect-uri>`.

    :::image type="content" source="../images/azure-portal-register-an-application.png" alt-text="Registre um painel de aplicativo com o nome e a conta com suporte concluída.":::

1. Selecione **Registrar**. Uma mensagem é exibida informando que o registro do aplicativo foi criado.

    :::image type="content" source="../images/azure-portal-application-created-message.png" alt-text="Mensagem informando que o registro do aplicativo foi criado.":::

1. Copie e salve os valores da **ID do Aplicativo (cliente)** e da **ID do Diretório (locatário).** Use ambos os valores nos procedimentos posteriores.

    :::image type="content" source="../images/azure-portal-copy-client-directory-ids.png" alt-text="Painel de registro de aplicativo da Contoso que exibe a ID do cliente e a ID do diretório.":::

## <a name="add-a-client-secret"></a>Adicionar um segredo do cliente

Às vezes chamado de _senha de aplicativo_, um segredo do cliente é um valor de cadeia de caracteres que seu aplicativo pode usar no lugar de um certificado para se identificar.

1. Selecione **Certificados & segredos**. Em seguida, na guia **Segredos do cliente** , selecione **Novo segredo do cliente**.

    :::image type="content" source="../images/azure-portal-create-new-client-secret.png" alt-text="O painel Certificados & segredos.":::

    O painel **Adicionar um segredo do cliente** é exibido.

1. Adicione uma descrição para o segredo do cliente.
1. Selecione uma expiração para o segredo ou especifique um tempo de vida personalizado.
    * O tempo de vida do segredo do cliente é limitado a dois anos (24 meses) ou menos. Você não pode especificar uma vida útil personalizada com mais de 24 meses.
    * A Microsoft recomenda que você defina um valor de expiração inferior a 12 meses.

    :::image type="content" source="../images/azure-portal-client-secret-description.png" alt-text="Adicione um painel de segredo do cliente com a descrição e expira concluído.":::

1. Selecione **Adicionar**. O novo segredo é criado, o valor é exibido temporariamente.

> [!IMPORTANT]
> _Registre o valor do segredo_ para uso no código do aplicativo cliente. Esse valor secreto _nunca será exibido novamente_ depois que você sair deste painel.

## <a name="expose-a-web-api"></a>Expor uma API Web

1. Selecione **Expor uma API**.

    O painel **Expor uma API** é exibido.

    :::image type="content" source="../images/azure-portal-expose-an-api.png" alt-text="Um painel Expor uma API de um registro de aplicativo.":::

1. Selecione **Definir** para gerar um URI de ID do aplicativo.

    :::image type="content" source="../images/azure-portal-set-api-uri.png" alt-text="Defina o botão no painel Expor uma API do registro do aplicativo.":::

    A seção para definir o URI de ID do aplicativo é exibida com um URI de ID de Aplicativo gerado no formulário `api://<app-id>`.

1. Atualize o URI da ID do aplicativo para `api://localhost:44355/<app-id>`.

    :::image type="content" source="../images/azure-portal-app-id-uri-details.png" alt-text="Edite o painel URI da ID do aplicativo com a porta localhost definida como 44355.":::

    * O **URI da ID do aplicativo** está preenchido previamente com a ID do aplicativo (GUID) no formato `api://<app-id>`.
    * O formato URI da ID do aplicativo deve ser: `api://<fully-qualified-domain-name>/<app-id>`
    * Insira o `fully-qualified-domain-name` entre `api://` e `<app-id>` (que é um GUID). Por exemplo, `api://contoso.com/<app-id>`.
    * Se você estiver usando localhost, o formato deverá ser `api://localhost:<port>/<app-id>`. Por exemplo, `api://localhost:44355/c6c1f32b-5e55-4997-881a-753cc1d563b7`.

    Para obter detalhes adicionais do URI da ID do aplicativo, consulte [Atributo identifierUris do manifesto do aplicativo](/azure/active-directory/develop/reference-app-manifest#identifieruris-attribute).

    > [!NOTE]
    > Se você receber um erro dizendo que o domínio já pertence a alguém, mas você é o seu proprietário, siga o procedimento em [Início Rápido: Adicionar um domínio personalizado ao Azure Active Directory](/azure/active-directory/add-custom-domain) para registrá-lo e, em seguida, repita esta etapa. (Esse erro também pode ocorrer se você não estiver conectado com credenciais de um administrador no locatário do Microsoft 365. (Confira a etapa 2.) Saia e entre novamente com credenciais de administrador e repita o processo da etapa 3.)

## <a name="add-a-scope"></a>Adicionar um escopo

1. Selecione **Adicionar um escopo**.

    :::image type="content" source="../images/azure-portal-add-a-scope.png" alt-text="Selecione Adicionar um botão de escopo.":::

    O painel **Adicionar um escopo** é aberto.

1. No painel **Adicionar um escopo** , especifique os atributos do escopo .

    :::image type="content" source="../images/azure-portal-add-a-scope-details.png" alt-text="Adicione um painel de escopo com valores de exemplo.":::

    | Campo | Descrição | Valores |
    |-------|-------------|---------|
    | **Nome do Escopo** | O nome do escopo. Uma convenção de nomenclatura de escopo comum é `resource.operation.constraint`. | Para SSO, isso deve ser definido como `access_as_user`. |
    | **Quem pode consentir?** |  Determina se o consentimento do administrador é necessário ou se os usuários podem consentir sem uma aprovação de administrador. | Para aprender SSO e exemplos, recomendamos que você defina isso como **administradores e usuários**. <br><br>Selecione **Administradores somente** para permissões com privilégios mais altos.|
    | **Administração nome de exibição de consentimento** | Uma breve descrição da finalidade do escopo visível apenas para administradores. | `Read-only access to user files and profiles.` |
    | **Administração descrição do consentimento** | Uma descrição mais detalhada da permissão concedida pelo escopo que somente os administradores veem. | `Allow Office to have read-only access to all user files and profiles. Office can call the app's web APIs as the current user.` |
    | **Nome de exibição de consentimento do usuário** | Uma breve descrição da finalidade do escopo. Mostrado aos usuários somente se você definir **Quem pode consentir** com **administradores e usuários**. | `Read-only access to your files and profile.` |
    | **Descrição do consentimento do usuário** | Uma descrição mais detalhada da permissão concedida pelo escopo. Mostrado aos usuários somente se você definir **Quem pode consentir** com **administradores e usuários**. | `Allow Office to have read-only access to your files and user profile.` |

1. Defina o **Estado** como **Habilitado** e selecione **Adicionar escopo**.

    :::image type="content" source="../images/azure-portal-enable-state-add-scope-button.png" alt-text="Defina o estado como habilitado e selecione o botão adicionar escopo.":::

    O novo escopo definido é exibido no painel.

    :::image type="content" source="../images/azure-portal-scope-added-successfully.png" alt-text="O novo escopo exibido no painel Expor uma API.":::

    > [!NOTE]
    > A parte de domínio do **Nome de escopo** exibidos logo abaixo do campo de texto deve corresponder automaticamente ao **URI de ID do aplicativo** definidos na etapa anterior com `/access_as_user` acrescentado ao final; por exemplo, `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`.

1. Selecione **Adicionar um aplicativo cliente**

    :::image type="content" source="../images/azure-portal-add-a-client-application.png" alt-text="Selecione adicionar um aplicativo cliente.":::

    O painel **Adicionar um aplicativo cliente** é exibido.

1. Na **ID do cliente, insira** `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e`. Esse valor pré-autoriza todos os pontos de extremidade do aplicativo do Microsoft Office.

    > [!NOTE]
    > A `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` ID pré-autoriza o Office em todas as plataformas a seguir. Como alternativa, você pode inserir um subconjunto adequado das seguintes IDs se, por qualquer motivo, quiser negar a autorização ao Office em algumas plataformas. Basta deixar de fora as IDs das plataformas das quais você deseja reter a autorização. Os usuários do suplemento nessas plataformas não poderão chamar suas APIs Web, mas outras funcionalidades no suplemento ainda funcionarão.
    >
    > - `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)
    > - `93d53678-613d-4013-afc1-62e9e444a0a5`(Office na Web)
    > - `bc59ab01-8403-45c6-8796-ac3ef710b3e3`(Outlook na Web)

1. em **Escopos autorizados**, selecione a caixa de seleção `api://localhost:44355/<app-id>/access_as_user` .

1. Selecione **Adicionar aplicativo**.

    :::image type="content" source="../images/azure-portal-add-application.png" alt-text="O painel Adicionar um aplicativo cliente.":::

## <a name="add-microsoft-graph-permissions"></a>Adicionar permissões do Microsoft Graph

1. Selecione **Permissões de API**.

    :::image type="content" source="../images/azure-portal-api-permissions.png" alt-text="O painel permissões de API.":::

    O painel **de permissões de API** é aberto.

1. Selecione **Adicionar uma permissão**.

    :::image type="content" source="../images/azure-portal-add-a-permission.png" alt-text="Adicionando uma permissão no painel de permissões de API.":::

    O painel **Solicitar permissões de API** é aberto.

1. Selecione **Microsoft Graph**.

    :::image type="content" source="../images/azure-portal-request-api-permissions-graph.png" alt-text="O painel Solicitar permissões de API com o botão Microsoft Graph.":::

1. Selecione **Permissões delegadas**.

    :::image type="content" source="../images/azure-portal-request-api-permissions-delegated.png" alt-text="O painel Solicitar permissões de API com o botão permissões delegadas.":::

1. Na caixa **de pesquisa Selecionar permissões, pesquise** as permissões que seu suplemento precisa. A seguir estão valores típicos usados nos exemplos.

    * Files.Read
    * openid
    * perfil

    > [!NOTE]
    > A permissão `User.Read` pode já estar listada por padrão. É uma boa prática solicitar apenas permissões necessárias, portanto, recomendamos que você desmarque a caixa para essa permissão se o suplemento realmente não precisar dela.

1. Selecione a caixa de seleção para cada permissão conforme ela aparece. Observe que as permissões não permanecerão visíveis na lista enquanto você seleciona cada uma delas. Depois de selecionar as permissões que seu suplemento precisa, selecione **Adicionar permissões**.

    :::image type="content" source="../images/azure-portal-request-api-permissions-add-permissions.png" alt-text="O painel Solicitar permissões de API com algumas permissões selecionadas.":::

## <a name="configure-access-token-version"></a>Configurar a versão do token de acesso

Você deve definir a versão do token de acesso aceitável para seu aplicativo. Essa configuração é feita no manifesto do aplicativo do Azure Active Directory.

### <a name="define-the-access-token-version"></a>Definir a versão do token de acesso

A versão do token de acesso poderá ser alterada se você escolher um tipo de conta diferente de **Contas em qualquer diretório organizacional (Qualquer diretório Azure AD – Multilocatário) e contas pessoais da Microsoft (por exemplo, Skype, Xbox)**. Use as etapas a seguir para garantir que a versão do token de acesso esteja correta para o uso do SSO do Office.

1. Selecione **Gerenciar** > **Manifesto** no painel esquerdo.

    :::image type="content" source="../images/azure-portal-manifest.png" alt-text="Selecione Manifesto do Azure.":::

    O manifesto do aplicativo do Azure Active Directory é exibido.

1. Insira **2** como o valor da propriedade `accessTokenAcceptedVersion`.

    :::image type="content" source="../images/azure-portal-manifest-token-version.png" alt-text="Valor para a versão do token de acesso aceito.":::

1. Selecione **Salvar**

    Uma mensagem é exibida no navegador informando que o manifesto foi atualizado com êxito.

    :::image type="content" source="../images/azure-portal-manifest-updated-message.png" alt-text="Mensagem atualizada de manifesto.":::

Parabéns! Você concluiu o registro do aplicativo para habilitar o SSO para seu suplemento do Office.
