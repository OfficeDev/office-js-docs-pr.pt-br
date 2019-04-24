
> [!NOTE]
> Este procedimento só é necessário durante a criação do suplemento. Quando seu suplemento de produção é implantado no AppSource ou em um catálogo de suplementos, os usuários confiarão individualmente nele ou um administrador se consentirá na organização na instalação.

Execute este procedimento *depois* [de registrar o suplemento](../develop/register-sso-add-in-aad-v2.md). (Se você acabou de concluir esse procedimento e a guia **permissões de API** da página **$Add-in-name $** estiver aberta no navegador, você pode escolher o botão **conceder consentimento de administrador para [nome do locatário]** e selecionar **Sim** para a confirmação que aparece. Pule o restante deste procedimento.)

1. Navegue até a página [Azure portal-app registrations](https://go.microsoft.com/fwlink/?linkid=2083908) para exibir o registro do aplicativo.

1. Entre com as credenciais de ***administrador*** em sua locação do Office 365. Por exemplo, MeuNome@contoso.onmicrosoft.com.

1. Selecione o aplicativo com o nome para exibição **$Add-in-name $**.

1. Na página **$Add-in-name $** , selecione **permissões de API** e, na seção **conceder consentimento** , escolha o botão **conceder consentimento de administrador para [nome do locatário]** . Selecione **Sim** para a confirmação exibida.

> [!NOTE]
> Recomendamos esse procedimento como prática recomendada se você estiver usando um locatário do O365 do desenvolvedor. No enTanto, se preferir, é possível Sideload um suplemento SSO em desenvolvimento e solicitar ao usuário um formulário de consentimento. Para obter mais informações, consulte [Sideload no Windows](/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins) e [Sideload no Office Online](/office/dev/add-ins/testing/sideload-office-add-ins-for-testing).
