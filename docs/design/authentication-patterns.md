# <a name="authentication-patterns"></a>Padrões de autenticação

Suplementos podem exigir que os usuários entrem ou se inscrevam para acessar recursos e funcionalidades. Caixas de entrada para nome de usuário e senha ou botões que iniciam fluxos de credenciais de terceiros são controles de interface comuns em experiências de autenticação. Uma experiência de autenticação simples e eficiente é um passo inicial importante para que os usuários comecem a interagir com seu suplemento.

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
> Atualmente a API de logon único tem suporte em versão prévia para Word, Excel e PowerPoint. Confira mais informações sobre os programas para os quais a API de logon único tem suporte no momento em [Conjuntos de requisitos da IdentityAPI](https://docs.microsoft.com/javascript/office/requirement-sets/identity-api-requirement-sets?view=office-js). Se você estiver trabalhando com um suplemento do Outlook, certifique-se de habilitar a Autenticação Moderna para a locação do Office 365. Confira mais informações sobre como fazer isso em [Exchange Online: como habilitar seu locatário para autenticação moderna](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).


## <a name="authentication-flow"></a>Fluxo de autenticação
Se o logon único ainda não estiver disponível para seus usuários, considere um fluxo de autenticação alternativo. Dê aos usuários a opção de se conectar diretamente com seu serviço ou com um provedor de identidade como a Microsoft.

1. Placemat da tela de apresentação: coloque o botão de entrar como uma solicitação de ação clara dentro da tela de apresentação do seu suplemento.
![](../images/add-in-fre-value-placemat.png)

2. Diálogo de Opções do Provedor de Identidade: exibe uma lista clara de provedores de identidade, incluindo um formulário de nome de usuário e senha, quando aplicável. A interface do usuário de seu suplemento pode ser bloqueada enquanto o diálogo de autenticação estiver aberto.
![](../images/add-in-auth-choices-dialog.png)



3. Conexão do Provedor de Identidade: o provedor de identidade terá sua própria interface do usuário. O Active Directory do Microsoft Azure permite a personalização de páginas de credenciais e de painel de acesso para uma aparência consistente com seu serviço. [Saiba mais](https://docs.microsoft.com/azure/active-directory/fundamentals/customize-branding).
![](../images/add-in-auth-identity-sign-in.png)

4. Andamento: indica o andamento enquanto as configurações e a interface do usuário são carregadas.
![](../images/add-in-auth-modal-interstitial.png)

> [!NOTE] 
> Ao usar o serviço de identidade da Microsoft, você terá a oportunidade de usar um botão personalizado de entrada que é ajustável para temas claros e escuros. Saiba mais.

## <a name="single-sign-on-authentication-flow"></a>Fluxo de autenticação de logon único
O logon único ainda é uma versão prévia. Assim que ele estiver disponível a todos, use-o para oferecer a melhor experiência ao usuário final. A identidade do usuário no Office é usada para se conectar ao seu suplemento. Desta forma, os usuários só se conectam uma vez. Isso deixa a experiência mais fluida, fazendo com que seja mais fácil de começar para seus clientes.

1. Ao instalar um suplemento, o usuário verá uma janela de consentimento semelhante a esta: ![](../images/add-in-auth-SSO-consent-dialog.png)
> [!NOTE]
> O publicador de suplementos terá controle sobre: o logotipo, as cadeias de caracteres e os escopos de permissão, incluídos na janela de consentimento. A interface do usuário é pré-configurada pela Microsoft.

2. O suplemento será carregado depois que o usuário der seu consentimento. É possível extrair do usuário e exibir a ele qualquer informação personalizada necessária.
![](../images/add-in-ribbon.png)

## <a name="see-also"></a>Confira também
- Saiba mais sobre [desenvolvimento de suplementos de logon único](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins)