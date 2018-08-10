
> [!NOTE]
> Este procedimento só é necessário durante a criação do suplemento. Quando o suplemento de produção for implantado no AppSource ou em um catálogo de suplementos, os usuários confiarão individualmente nele ou um administrador concordará pela organização na instalação.

Execute este procedimento *após* você ter [registrado o suplemento](../develop/register-sso-add-in-aad-v2.md).

1. Na sequência de caracteres a seguir, substitua o espaço reservado "{application_ID}" pela ID do aplicativo que você copiou quando registrou seu suplemento:  `https://login.microsoftonline.com/common/adminconsent?client_id={application_ID}&state=12345`

1. Cole a URL resultante na barra de endereços do navegador e acesse-a.

1. Quando for solicitado, entre com as credenciais de administrador em sua locação do Office 365.

1. Em seguida, será solicitado que você conceda permissão para seu suplemento acessar os dados do Microsoft Graph. Clique em **Aceitar**.

1. A janela/guia do navegador é redirecionada para a **URL de redirecionamento** que você especificou quando registrou o suplemento. Se o aplicativo da Web do suplemento estiver em execução, a home page do suplemento será aberta no navegador; caso contrário, você receberá um erro 404. Mas o fato de que o navegador tentou abrir a home page significa que o consentimento foi concedido com sucesso.

>[!NOTE]
>Recomendamos esse procedimento como uma melhor prática se você estiver usando um locatário do Developer O365. No entanto, se preferir, é possível carregar um suplemento do SSO em desenvolvimento e solicitar ao usuário um formulário de consentimento. Para mais informações, veja [Sideload no Windows](https://docs.microsoft.com/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins) e [Sideload no Office Online](https://docs.microsoft.com/office/dev/add-ins/testing/sideload-office-add-ins-for-testing).

