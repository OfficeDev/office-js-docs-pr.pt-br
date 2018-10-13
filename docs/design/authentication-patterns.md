# <a name="authentication-patterns"></a>Padrões de autenticação

Os suplementos podem exigir que os usuários façam login ou se inscrevam para acessar recursos e funcionalidades. As caixas de entrada para nome de usuário e senha ou os botões que iniciam fluxos de credenciais de terceiros são controles de interface comuns em experiências de autenticação. Uma experiência de autenticação simples e eficiente é um primeiro passo importante para que os usuários comecem a usar seu suplemento.

## <a name="best-practices"></a>Práticas recomendadas

|Fazer|Não fazer|
|:----|:----|
|Usar o logon único (SSO) para autenticar usuários em seu suplemento.|Exigir que usuários conectem-se ao seu suplemento usando credenciais diferentes de suas contas pessoais da Microsoft ou de suas contas do Office 365 (profissional ou estudantil).|
|Antes de fazer o usuário se conectar, descrever o valor de seu suplemento ou demonstrar funcionalidades sem exigir uma conta. |Esperar que os usuários se conectem sem entender o valor e os benefícios de seu complemento.|
|Guiar os usuários por meio de fluxos de autenticação com um botão primário, altamente visível em cada tela. |Chamar atenção para tarefas secundárias e terciárias com botões e solicitações de ação concorrentes entre si.|
|Usar rótulos claros de botão que descrevam tarefas específicas, como "Entrar" ou "Criar conta".   |Usar rótulos de botão vagos, como "Enviar" ou "Começar" para orientar os usuários em fluxos de autenticação.|
|Usar um diálogo para focar a atenção dos usuários nos formulários de autenticação.    |Sobrecarregar seu painel de tarefas com uma tela de apresentação e formulários de autenticação.|
|Encontrar pequenas eficiências no fluxo, como foco automático nas caixas de entrada. |Adicionar etapas desnecessárias à interação, como exigir que os usuários cliquem nos campos do formulário.|
|Fornecer uma maneira para os usuários saírem e se autenticarem novamente.    |Forçar usuários a desinstalar para alternar entre identidades.|

> [!NOTE]
> A API de logon único é atualmente compatível com as versões prévias do Word, Excel, Outlook e PowerPoint. Para mais informações sobre a compatibilidade da API de logon único, veja [Conjuntos de requisitos da API de identidade](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets?view=office-js). Se você estiver trabalhando com um suplemento do Outlook, não esqueça de habilitar a Autenticação Moderna para o locatário do Office 365. Para saber como fazer isso, veja [Exchange Online: Como habilitar o seu locatário para a Autenticação Moderna](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).


## <a name="authentication-flow"></a>Fluxo de autenticação
Se o logon único ainda não estiver disponível para seus usuários, considere um fluxo de autenticação alternativo. Dê aos usuários a opção de entrar diretamente com seu serviço ou um provedor de identidade como a Microsoft.

1. Placemat da tela de apresentação: coloque o botão de login como uma solicitação de ação clara dentro da tela de apresentação do seu suplemento.
![](../images/add-in-fre-value-placemat.png)

2. Caixa de diálogo de opções de provedores de identidade - exiba uma lista clara de provedores de identidade, incluindo um formulário de nome de usuário e senha, se aplicável. A interface do usuário do suplemento pode ser bloqueada enquanto a caixa de diálogo de autenticação estiver aberta. ![](../images/add-in-auth-choices-dialog.png)



3. Login do provedor de identidade - o provedor de identidade terá sua própria interface do usuário. O Active Directory do Microsoft Azure permite a personalização de páginas de login e painel de acesso para uma aparência consistente com seu serviço. [Saiba mais](https://docs.microsoft.com/azure/active-directory/fundamentals/customize-branding). ![](../images/add-in-auth-identity-sign-in.png)

4. Progresso - indica o progresso enquanto as configurações e o a interface do usuário são carregadas.
![](../images/add-in-auth-modal-interstitial.png)

> [!NOTE] 
> Ao usar o serviço de identidade da Microsoft, você terá a oportunidade de usar um botão personalizado de login que é ajustável para temas claros e escuros. Saiba mais.

## <a name="single-sign-on-authentication-flow"></a>Fluxo de autenticação de logon único
O logon único ainda está em versão prévia. Uma vez disponível para todos, use-o para que o usuário final tenha uma melhor experiência. A identidade do usuário no Office é usada para fazer login no seu suplemento. Como resultado, os usuários só fazem login uma vez. Isso elimina as dificuldades na experiência, facilitando o uso de seus clientes.

1. Ao instalar um suplemento, o usuário verá uma janela de consentimento semelhante a esta: ![](../images/add-in-auth-SSO-consent-dialog.png)
> [!NOTE]
> O publicador de suplementos terá controle sobre o logotipo, as sequências de caracteres e os escopos de permissão, incluídos na janela de consentimento. A interface do usuário é pré-configurada pela Microsoft.

2. O suplemento será carregado depois que o usuário consentir. Ele pode extrair e exibir qualquer informação personalizada necessária ao usuário. ![](../images/add-in-ribbon.png)

## <a name="see-also"></a>Confira também
- Saiba mais sobre [desenvolvimento de suplementos de logon único](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins)