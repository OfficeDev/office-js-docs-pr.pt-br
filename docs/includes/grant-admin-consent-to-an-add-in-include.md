
> [!NOTE]
> Este procedimento só é necessário durante a criação do suplemento. Quando o seu complemento de produção é implantado no AppSource ou em um catálogo de aplicativos, os usuários confiarão individualmente nele ou um administrador consentiria com a organização na instalação.

Realize este procedimento *depois de* ter registrado [o add-in](../develop/register-sso-add-in-aad-v2.md). (Se você tiver concluído esse procedimento e a guia permissões de **API** da página **$ADD-IN-NAME$** estiver aberta no navegador, você poderá escolher o botão Conceder consentimento de administrador **para [nome** do locatário] e, em seguida, selecione **Sim** para a confirmação exibida. Ignore o restante deste procedimento.)

1. Navegue até [o portal do Azure - Página de registros de aplicativos](https://go.microsoft.com/fwlink/?linkid=2083908) para exibir o registro do aplicativo.

1. Entre com as ***credenciais de*** administrador no seu Microsoft 365 de adoção. Por exemplo, MeuNome@contoso.onmicrosoft.com.

1. Selecione o aplicativo com nome para **exibição $ADD-IN-NAME$**.

1. Na página **$ADD-IN-NAME$,** selecione permissões de **API,**  em seguida, na seção Conceder consentimento, escolha o botão Conceder consentimento de administrador **para [nome** do locatário]. Selecione **Sim** para a confirmação exibida.

> [!NOTE]
> Recomendamos este procedimento como uma prática prática prática se você estiver usando um locatário do Developer O365. No entanto, se preferir, é possível fazer sideload de um complemento SSO em desenvolvimento e solicitar ao usuário um formulário de consentimento. Para obter mais informações, [consulte Sideload on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) and [Sideload on Office na Web](../testing/sideload-office-add-ins-for-testing.md).
