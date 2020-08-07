

1. Acesse a página [Portal do Azure - Registros de aplicativo](https://go.microsoft.com/fwlink/?linkid=2083908) para registrar o seu aplicativo.

1. Entre com as credenciais de ***administrador*** em seu Microsoft 365 locação. Por exemplo, MeuNome@contoso.onmicrosoft.com.

1. Selecione **Novo registro**. Na página **Registrar um aplicativo**, defina os valores da seguinte forma.

    * Definir **Nome** ao **$SUPLEMENTO-NOME$**.
    * Defina **Tipos de conta com suporte** para **Contas em qualquer diretório organizacional e contas pessoais da Microsoft (por exemplo, Skype, Xbox, Outlook.com)**.
    * Deixe o **URI de Redirecionamento** vazio.
    * Escolha **Registrar**.

1. Na página **$SUPLEMENTO-NOME$**, copie e salve os valores para a **ID do aplicativo (cliente)** e a **ID do diretório (locatário)**. Use ambos os valores nos procedimentos posteriores.

    > [!NOTE]
    > Essa ID é o valor "audience" (público) quando outros aplicativos, como o aplicativo host do Office (por exemplo, PowerPoint, Word, Excel), buscam o acesso autorizado ao aplicativo. Também é a "ID do cliente" do aplicativo quando ela, por sua vez, busca acesso autorizado ao Microsoft Graph.

1. Selecione **Certificados e segredos** sob **Gerenciar**. Selecione o botão **Novo segredo do cliente**. Insira um valor para **Descrição** e, em seguida, selecione uma opção adequada para **Expira** e escolha **Adicionar**. *Copiar o valor de segredo do cliente imediatamente e salvá-lo com a ID de aplicativo* antes de prosseguir, pois ele será necessário em um procedimento posterior.

1. Selecionar **Expor uma API** em **Gerenciar**. Selecione o link **Definir** para gerar o URI da ID de Aplicativo no formato "api: / / $App ID GUID$". Inserir o **$FQDN-WITHOUT-PROTOCOL$** (com uma barra "/" acrescentada ao final) entre as duas barras e o GUID. A ID inteira deve ter o formulário `api://$FQDN-WITHOUT-PROTOCOL$/$App ID GUID$`; por exemplo `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7`.

    > [!NOTE]
    > Você pode receber um erro impreciso nesse momento, dizendo "O URI da ID de aplicativo deve ser um URI válido começando HTTPS, API, URN, MS APPX. Ele não pode terminar com uma barra." Se a ID atende às condições mencionadas, ignore o erro e salve suas alterações.

    > [!NOTE]
    > Se você receber um erro dizendo que o domínio já pertence a alguém, mas você é o seu proprietário, siga o procedimento em [Início Rápido: Adicionar um domínio personalizado ao Azure Active Directory](/azure/active-directory/add-custom-domain) para registrá-lo e, em seguida, repita esta etapa. (Esse erro também pode ocorrer se você não tiver entrado com as credenciais de um administrador no Microsoft 365 locação. (Confira a etapa 2.) Saia e entre novamente com credenciais de administrador e repita o processo da etapa 3.)

1. Selecione o botão **Adicionar um escopo**. No painel que se abre, insira `access_as_user` como o **Nome de escopo**.

1. Definir **Quem pode consentir?** aos **Administradores e usuários**.

1. Preencha os campos para configurar a solicitação de consentimento de administrador e usuário com valores apropriados ao `access_as_user` escopo que permite que o aplicativo de host do Office use os seus APIs de suplemento da web com os mesmos direitos que o usuário atual. Sugestões:

    - **Título de autorização de administrador:** Office pode funcionar como o usuário.
    - **Descrição de autorização de administrador:** Permite ao Office chamar os APIs de suplemento da web com os mesmos direitos que o usuário atual.
    - **Título de autorização de usuário:** O Office pode funcionar como se fosse você.
    - **Descrição de autorização de administrador:** Permite ao Office chamar os APIs de suplemento da web com os mesmos direitos que você possui.

1. Verifique se o **Estado** está definido como **Habilitado**.

1. Selecione **Adicionar escopo**.

    > [!NOTE]
    > A parte de domínio do **Nome de escopo** exibidos logo abaixo do campo de texto deve corresponder automaticamente ao **URI de ID do aplicativo** definidos na etapa anterior com `/access_as_user` acrescentado ao final; por exemplo, `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`.

1. Na seção **Aplicativos clientes autorizados**, você identifica os aplicativos que deseja autorizar para o aplicativo da Web do seu suplemento. Cada uma das seguintes IDs precisa ser pré-autorizada.
  
    * `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)
    * `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Microsoft Office)
    * `57fb890c-0dab-4253-a5e0-7188c88b2bb4`(Office na Web)
    * `08e18876-6177-487e-b8b5-cf950c1e598c`(Office na Web)
    * `bc59ab01-8403-45c6-8796-ac3ef710b3e3`(Outlook na Web)

    Para cada ID, siga estas etapas:

      a. Selecione o botão **Adicionar um aplicativo cliente** e, no painel que se abre, defina o **ID do cliente** para o respectivo GUID e marque a caixa `api://$FQDN-WITHOUT-PROTOCOL$/$App ID GUID$/access_as_user`.

      b. Selecione **Adicionar aplicativo**.

1. Selecione **Autenticação** em **Gerenciar**. Na seção **URIs de redirecionamento**, selecione **Web** no **Tipo** de lista suspensa, em seguida, defina o valor do**URI de redirecionamento** para `https://$FQDN-WITHOUT-PROTOCOL$`.

1. Na parte superior da página, selecione **Salvar**.

1. Selecione **Permissões para API** em **Gerenciar** e selecione **Adicionar uma permissão**. No painel que se abre, escolha **Microsoft Graph** e, em seguida, escolha **Permissões delegadas**.

1. Use a caixa de pesquisa **Selecionar permissões** para procurar as permissões que o seu suplemento precisa. Eis alguns exemplos.

    * Files.Read.All
    * offline_access
    * openid
    * perfil

    > [!NOTE]
    > A permissão `User.Read` pode já estar listada por padrão. É uma boa prática não pedir permissões desnecessárias, por isso recomendamos desmarcar a caixa para essa permissão se o suplemento não precisar dela.

1. Marque a caixa de seleção para cada permissão como aparece (observe que as permissões não permanecem visíveis na lista ao selecionar cada uma delas). Depois de selecionar as permissões que o suplemento precisa, selecione o botão **Adicionar permissões** na parte inferior do painel.
