# <a name="authentication-patterns"></a>Padrões de autenticação

Os suplementos podem exigir que os usuários entrem ou se inscrevam para acessar os recursos e funcionalidades. As caixas de entrada de nome de usuário e senha ou botões que iniciam fluxos de credenciais de terceiros são controles comuns da interface em experiências de autenticação. Uma experiência de autenticação simples e eficiente é uma primeira etapa importante para iniciar os usuários no uso de seu suplemento.

## <a name="best-practices"></a>Práticas recomendadas

|Fazer|Não fazer|
|:----|:----|
|Antes de entrar, descreva o valor do suplemento ou demonstre a funcionalidade sem exigir uma conta. |Espere que os usuários entrem sem compreender o valor e os benefícios do suplemento.|
|Oriente os usuários pelos fluxos de autenticação com um botão principal bem visível em cada tela. |Chame atenção para as tarefas secundárias e terciárias com outros botões e chamadas para ação.|
|Use rótulos de botão claros que descrevam tarefas específicas, como “Entrar” ou “Criar conta”.   |Use rótulos de botão vagos como “Enviar” ou “Começar” para orientar os usuários por meio de fluxos de autenticação.|
|Use uma caixa de diálogo para concentrar a atenção do usuário em formulários de autenticação.    |Encha seu painel de tarefas com uma primeira experiência de execução e formulários de autenticação.|
|Inclua pequenos recursos eficientes no fluxo como foco automático em caixas de entrada. |Adicione etapas desnecessárias à interação como exigir que os usuários cliquem nos campos de formulário.|
|Ofereça uma maneira para os usuários saírem e autenticarem-se novamente.    |Force os usuários a fazer a desinstalação para alternar identidades.|

## <a name="authentication-flow"></a>Fluxo de autenticação
Até o logon único estar fora da versão prévia, os suplementos de produção devem conceder aos usuários uma opção para entrar diretamente com o serviço ou um provedor de identidade como a Microsoft.

1. Marcador de primeira execução: coloque o botão de entrada como uma chamada para ação clara na primeira experiência de execução do seu suplemento.
![](../images/add-in-fre-value-placemat.png)

2. Caixa de diálogo de opções do provedor de identidade: exibe uma lista clara de provedores de identidade incluindo um formulário de nome de usuário e senha, se aplicável. A interface de usuário de um suplemento poderá ser bloqueada quando a caixa de diálogo de autenticação estiver aberta.
![](../images/add-in-auth-choices-dialog.png)



3. Entrada de um provedor de identidade: os provedores de identidade têm suas próprias interfaces de usuário. O Microsoft Azure Active Directory permite a personalização das páginas de entrada e de painel de acesso para uma aparência consistente com o serviço.  [Saiba mais](https://docs.microsoft.com/azure/active-directory/fundamentals/customize-branding).
![](../images/add-in-auth-identity-sign-in.png)

4. Progresso: indica o progresso enquanto as configurações e a interface do usuário são carregadas.
![](../images/add-in-auth-modal-interstitial.png)

> [!NOTE] 
> Ao usar o serviço de identidade da Microsoft, você terá a oportunidade de usar um botão de entrada com marca que poderá ser personalizado com temas claros e escuros.Saiba mais.

## <a name="single-sign-on-authentication-flow-preview"></a>Fluxo de autenticação de logon único (versão prévia)

> [!NOTE]
> Atualmente a API de logon único tem suporte na visualização para Word, Excel, Outlook e PowerPoint. Para saber mais informações sobre o suporte a logon único, confira  [Conjuntos de requisitos da IdentityAPI](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets?view=office-js). Se você estiver trabalhando com um suplemento do Outlook, certifique-se de habilitar a Autenticação Moderna para a locação do Office 365. Para saber mais informações sobre como fazer isso, confira  [Exchange Online: como habilitar seu locatário para autenticação moderna](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

Depois que o recurso de logon único for disponibilizado para suplementos de produção, use-o em uma experiência de usuário final mais estável. A identidade do usuário no Office (identidade da conta Microsoft ou do Office 365) é usada para entrar no suplemento. Como resultado, os usuários somente entram uma vez. Isso remove conflitos na experiência e faz com que os clientes comecem a usar o suplemento sem dificuldades.

1. Conforme um suplemento está sendo instalado, um usuário verá uma janela de consentimento semelhante ao exemplo abaixo: ![](../images/add-in-auth-SSO-consent-dialog.png)
> [!NOTE]
> O fornecedor do suplemento terá controle sobre o logotipo, sobre as cadeias de caracteres e escopos de permissão incluídos na janela de consentimento. A interface do usuário é pré-configurada pela Microsoft.

2. O suplemento será carregado após o consentimento do usuário. Ele pode extrair e exibir quaisquer informações de usuário personalizadas necessárias.
![](../images/add-in-ribbon.png)

## <a name="see-also"></a>Confira também
- Saiba mais sobre como [desenvolver suplementos de SSO (versão prévia)](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins)