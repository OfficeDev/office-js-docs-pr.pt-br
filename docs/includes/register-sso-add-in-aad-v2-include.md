## <a name="create-an-app-registration"></a>Criar um registro de aplicativo

Registrar seu aplicativo (o suplemento) estabelece uma relação de confiança entre o suplemento e o plataforma de identidade da Microsoft. A relação de confiança é unidirecional: seu suplemento confia no plataforma de identidade da Microsoft, e não o contrário.

1. Entre [no portal do Azure com](https://portal.azure.com/) as credenciais ***admin** _ para sua Microsoft 365 locação. Por exemplo, _*MyName@contoso.onmicrosoft.com**.
1. Em **Gerenciar**, selecione **Registros de aplicativo** >  **Novo registro**. Na página **Registrar um aplicativo**, defina os valores da seguinte forma.

    * Defina **Nome** para `<add-in-name>`.
    * **Defina os tipos** de conta com suporte como Contas em qualquer diretório organizacional (qualquer diretório do **Azure AD – multilocatário) e contas pessoais da Microsoft (por exemplo, Skype, Xbox)**.
    * Deixe o **URI de Redirecionamento** vazio.
    * Escolha **Registrar**.

1. Copie e salve os valores para a **ID do Aplicativo (cliente)** e a **ID do Diretório (locatário**). Use ambos os valores nos procedimentos posteriores.

    > [!NOTE]
    > Essa ID é o valor de "audiência" quando outros aplicativos, como o aplicativo cliente Office (por exemplo, PowerPoint, Word, Excel), buscam acesso autorizado ao aplicativo. Também é a "ID do cliente" do aplicativo quando ela, por sua vez, busca acesso autorizado ao Microsoft Graph.

## <a name="add-a-client-secret"></a>Adicionar um segredo do cliente

Às vezes chamado de _senha de aplicativo_, um segredo do cliente é um valor de cadeia de caracteres que seu aplicativo pode usar no lugar de um certificado para se identificar.

1. No portal do Azure, em **Registros de aplicativo**, selecione seu aplicativo.
1. Selecione **Certificados &** **secretsClient** >  secretsNew  > **segredo do cliente**.
1. Adicione uma descrição para o segredo do cliente.
1. Selecione uma expiração para o segredo ou especifique um tempo de vida personalizado.
    * O tempo de vida do segredo do cliente é limitado a dois anos (24 meses) ou menos. Você não pode especificar um tempo de vida personalizado por mais de 24 meses.
    * A Microsoft recomenda que você defina um valor de expiração inferior a 12 meses.
1. Selecione **Adicionar**.
1. _Registre o valor do segredo para_ uso no código do aplicativo cliente. Esse valor secreto nunca _será exibido novamente depois_ que você sair desta página.

## <a name="expose-a-web-api"></a>Expor uma API Web

1. Verifique se você está exibindo o registro de aplicativo que acabou de criar.
1. Em **Gerenciar**, selecione **Expor uma API** e **selecione o link** Definir. Isso abre uma **caixa Definir o URI da ID** do Aplicativo com um URI de ID do Aplicativo gerado no formulário `api://<application-id>`. Insira seu nome de domínio totalmente qualificado antes de `<application-id>`. A ID inteira deve ter o formulário `api://<fully-qualified-domain-name>/<application-id>`; por exemplo, `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7`.

    > [!NOTE]
    > Se você receber um erro dizendo que o domínio já pertence a alguém, mas você é o seu proprietário, siga o procedimento em [Início Rápido: Adicionar um domínio personalizado ao Azure Active Directory](/azure/active-directory/add-custom-domain) para registrá-lo e, em seguida, repita esta etapa. (Esse erro também poderá ocorrer se você não estiver conectado com as credenciais de um administrador no Microsoft 365 locatário. (Confira a etapa 2.) Saia e entre novamente com credenciais de administrador e repita o processo da etapa 3.)

## <a name="add-a-scope"></a>Adicionar um escopo

1. Selecione o botão **Adicionar um escopo**. No painel que se abre, insira `access_as_user` como o **Nome de escopo**.

1. Definir **Quem pode consentir?** aos **Administradores e usuários**.

1. Preencha os campos para configurar os prompts de consentimento do administrador e do usuário com valores apropriados `access_as_user` para o escopo que permite que o aplicativo cliente do Office use as APIs Web do suplemento com os mesmos direitos que o usuário atual. Sugestões:

    * **Nome de exibição do consentimento** do administrador: Office pode atuar como o usuário.
    * **Descrição de autorização de administrador:** Permite ao Office chamar os APIs de suplemento da web com os mesmos direitos que o usuário atual.
    * **Nome de exibição de** consentimento do usuário: Office pode agir como você.
    * **Descrição de autorização de usuário:** Permite ao Office chamar os APIs de suplemento da web com os mesmos direitos que você possui.

1. Verifique se o **Estado** está definido como **Habilitado**.

1. Selecione **Adicionar escopo**.

    > [!NOTE]
    > A parte de domínio do **Nome de escopo** exibidos logo abaixo do campo de texto deve corresponder automaticamente ao **URI de ID do aplicativo** definidos na etapa anterior com `/access_as_user` acrescentado ao final; por exemplo, `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`.

1. Na seção **Aplicativos cliente autorizados**, insira a ID a seguir para pré-autorizar todos os Microsoft Office de extremidade do aplicativo.

   - `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e`(Todos os Microsoft Office de extremidade do aplicativo)

    > [!NOTE]
    > A `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` ID pré-autoriza Office em todas as plataformas a seguir. Como alternativa, você pode inserir um subconjunto adequado das IDs a seguir se, por algum motivo, quiser negar a autorização para Office em algumas plataformas. Basta deixar de fora as IDs das plataformas das quais você deseja reprisar a autorização. Os usuários do suplemento nessas plataformas não poderão chamar suas APIs Web, mas outras funcionalidades no suplemento ainda funcionarão.
    >
    > - `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)
    > - `93d53678-613d-4013-afc1-62e9e444a0a5`(Office na Web)
    > - `bc59ab01-8403-45c6-8796-ac3ef710b3e3`(Outlook na Web)

1. Selecione **Adicionar um aplicativo cliente**. No painel que é aberto, defina a **ID do** Cliente como o respectivo GUID e marque a caixa para `api://<fully-qualified-domain-name>/<application-id>/access_as_user`.

1. Selecione **Adicionar aplicativo**.

## <a name="add-microsoft-graph-permissions"></a>Adicionar permissões Graph Microsoft

1. Em **Gerenciar**, selecione **Autenticação** e, em seguida, **escolha Adicionar uma plataforma**.

1. No painel **Configurar plataformas** , selecione **Web** e defina o valor do **URI de** Redirecionamento como `https://<fully-qualified-domain-name>`.

1. Escolha **Configurar**.

1. Em **Gerenciar**, selecione **permissões de API** e **selecione Adicionar uma permissão**. No painel que é aberto, escolha **Microsoft Graph** e, em seguida, escolha **Permissões delegadas**.

1. Use a caixa de pesquisa **Selecionar permissões** para procurar as permissões que o seu suplemento precisa. Eis alguns exemplos.

    * Files.Read.All
    * offline_access
    * openid
    * perfil

    > [!NOTE]
    > A permissão `User.Read` pode já estar listada por padrão. É uma boa prática solicitar apenas as permissões necessárias, portanto, recomendamos desmarcar a caixa dessa permissão se o suplemento não precisar dela.

1. Marque a caixa de seleção para cada permissão como aparece (observe que as permissões não permanecem visíveis na lista ao selecionar cada uma delas). Depois de selecionar as permissões de que seu suplemento precisa, selecione o **botão Adicionar permissões** .
