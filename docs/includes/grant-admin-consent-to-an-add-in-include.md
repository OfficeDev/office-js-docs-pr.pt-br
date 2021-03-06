
> [!NOTE]
> Este procedimento só é necessário durante a criação do suplemento. Quando seu suplemento de produção é implantado no AppSource ou em um catálogo de aplicativos, os usuários confiarão individualmente nele ou um administrador se consentirá na organização na instalação.

Execute este procedimento *depois* [de registrar o suplemento](../develop/register-sso-add-in-aad-v2.md). (Se você acabou de concluir esse procedimento e a guia **permissões de API** da página **$Add-in-name $** estiver aberta no navegador, você pode escolher o botão **conceder consentimento de administrador para [nome do locatário]** e, em seguida, selecione **Sim** para a confirmação exibida. Pule o restante deste procedimento.)

1. Navegue até a página [Azure portal-app registrations](https://go.microsoft.com/fwlink/?linkid=2083908) para exibir o registro do aplicativo.

1. Entre com as credenciais de ***administrador*** em seu Microsoft 365 locação. Por exemplo, MeuNome@contoso.onmicrosoft.com.

1. Selecione o aplicativo com o nome para exibição **$Add-in-name $**.

1. Na página **$Add-in-name $** , selecione **permissões de API** e, na seção **conceder consentimento** , escolha o botão **conceder consentimento de administrador para [nome do locatário]** . Selecione **Sim** para a confirmação exibida.

> [!NOTE]
> Recomendamos esse procedimento como prática recomendada se você estiver usando um locatário do O365 do desenvolvedor. No entanto, se preferir, é possível Sideload um suplemento SSO em desenvolvimento e solicitar ao usuário um formulário de consentimento. Para obter mais informações, consulte [Sideload no Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) e [Sideload no Office na Web](../testing/sideload-office-add-ins-for-testing.md).
