---
title: Opções de listagem do AppSource para seu suplemento do Outlook baseado em evento
description: Saiba mais sobre as opções de listagem do AppSource disponíveis para seu suplemento do Outlook que implementa a ativação baseada em eventos.
ms.topic: article
ms.date: 07/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: d8d2c2e9960d2aef2d32ede6e20eb5f1db125a6c
ms.sourcegitcommit: 9bb790f6264f7206396b32a677a9133ab4854d4e
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/15/2022
ms.locfileid: "66797677"
---
# <a name="appsource-listing-options-for-your-event-based-outlook-add-in"></a>Opções de listagem do AppSource para seu suplemento do Outlook baseado em evento

No momento, os suplementos devem ser implantados pelos administradores de uma organização para que os usuários finais acessem a funcionalidade de recurso baseada em evento. Estamos restringindo a ativação baseada em evento se o usuário final adquiriu o suplemento diretamente do AppSource. Por exemplo, se o suplemento Contoso `LaunchEvent` `LaunchEvent Type` `LaunchEvents` incluir o ponto de extensão com pelo menos um definido no nó, a invocação automática do suplemento só ocorrerá se o suplemento tiver sido instalado para o usuário final pelo administrador da organização. Caso contrário, a invocação automática do suplemento será bloqueada. Consulte o trecho a seguir de um manifesto de suplemento de exemplo.

```xml
...
<ExtensionPoint xsi:type="LaunchEvent">
  <LaunchEvents>
    <LaunchEvent Type="OnNewMessageCompose" FunctionName="onMessageComposeHandler"/>
    ...
```

Um usuário final ou administrador pode adquirir suplementos por meio do AppSource ou da Office Store no aplicativo. Se o cenário ou fluxo de trabalho principal do suplemento exigir a ativação baseada em evento, talvez você queira restringir os suplementos disponíveis para a implantação do administrador. Para habilitar essa restrição, podemos fornecer URLs de código de voo. Graças aos códigos de voo, somente os usuários finais com essas URLs especiais podem acessar a listagem. A seguir está um exemplo de URL.

`https://appsource.microsoft.com/product/office/WA200002862?flightCodes=EventBasedTest1`

Os usuários e administradores não podem pesquisar explicitamente um suplemento pelo nome no AppSource ou na Office Store no aplicativo quando um código de versão de pré-lançamento está habilitado para ele. Como criador do suplemento, você pode compartilhar esses códigos de versão de pré-lançamento de modo privado com os administradores da organização para implantação de suplementos.

> [!NOTE]
> Embora os usuários finais possam instalar o suplemento usando um código de versão de pré-lançamento, o suplemento não incluirá a ativação baseada em evento.

## <a name="specify-a-flight-code"></a>Especificar um código de versão de pré-lançamento

Para especificar o código de versão de pré-lançamento desejado para o suplemento, compartilhe essas informações  nas Notas para certificação quando estiver publicando o suplemento. _**Importante**:_ Os códigos de voo diferenciam maiúsculas de minúsculas.

![Captura de tela mostrando a solicitação de exemplo de código de voo em Anotações para a tela de certificação durante o processo de publicação.](../images/outlook-publish-notes-for-certification-1.png)

## <a name="deploy-add-in-with-flight-code"></a>Implantar suplemento com código de versão de pré-lançamento

Depois que os códigos de voo forem definidos, você receberá a URL da equipe de certificação do aplicativo. Em seguida, você pode compartilhar a URL com os administradores em particular.

Para implantar o suplemento, o administrador pode usar as etapas a seguir.

- Entre no admin.microsoft.com ou AppSource.com com sua conta de administrador do Microsoft 365. Se o suplemento tiver o SSO (logon único) habilitado, serão necessárias credenciais de administrador global.
- Abra a URL do código de voo em um navegador da Web.
- Na página de listagem de suplementos, selecione **Obter agora**. Você deve ser redirecionado para o portal do aplicativo integrado.

## <a name="unrestricted-appsource-listing"></a>Listagem irrestrita do AppSource

Se o suplemento não usar a ativação baseada em evento para cenários críticos (ou seja, seu suplemento funciona bem sem invocação automática), considere listar seu suplemento no AppSource sem nenhum código de voo especial. Se um usuário final receber seu suplemento do AppSource, a ativação automática não acontecerá para o usuário. No entanto, eles podem usar outros componentes do suplemento, como um painel de tarefas ou um comando de função.

> [!IMPORTANT]
> Essa é uma restrição temporária. No futuro, planejamos habilitar a ativação de suplemento baseado em evento para usuários finais que adquirem diretamente seu suplemento.

## <a name="update-existing-add-ins-to-include-event-based-activation"></a>Atualizar suplementos existentes para incluir a ativação baseada em evento

Você pode atualizar seu suplemento existente para incluir a ativação baseada em evento e, em seguida, reenviar para validação e decidir se deseja uma listagem restrita ou irrestrita do AppSource.

Depois que o suplemento atualizado for aprovado, os administradores da organização que já implantaram o suplemento receberão uma mensagem de atualização na seção Aplicativos integrados do centro de administração. A mensagem aconselha o administrador sobre as alterações de ativação baseadas em evento. Depois que o administrador aceitar as alterações, a atualização será implantada para os usuários finais.

![Captura de tela da notificação de atualização de aplicativo na tela "Aplicativos integrados".](../images/outlook-deploy-update-notification.png)

Para usuários finais que instalaram o suplemento por conta própria, o recurso de ativação baseada em evento não funcionará mesmo depois que o suplemento for atualizado.

## <a name="admin-consent-for-installing-event-based-add-ins"></a>Administração consentimento para instalar suplementos baseados em eventos

Sempre que um suplemento baseado em evento é implantado na tela Aplicativos Integrados, o administrador obtém detalhes sobre as funcionalidades de ativação baseada em eventos do suplemento no assistente de implantação. Os detalhes aparecem na seção **Permissões e Funcionalidades do** Aplicativo. O administrador deve ver todos os eventos em que o suplemento pode ser ativado automaticamente.

![Captura de tela da tela "Aceitar solicitações de permissões" ao implantar um novo aplicativo.](../images/outlook-deploy-accept-permissions-requests.png)

Da mesma forma, quando um suplemento existente é atualizado para a funcionalidade baseada em evento, o administrador vê um status de "Atualização Pendente" no suplemento. O suplemento atualizado será implantado somente se o administrador consentir com as alterações notadas na seção Permissões e **Funcionalidades** do Aplicativo, incluindo o conjunto de eventos em que o suplemento pode ser ativado automaticamente.

Sempre que você adicionar qualquer novidade `LaunchEvent Type` ao suplemento, os administradores verão o fluxo de atualização no portal de administração e precisarão fornecer consentimento para eventos adicionais.

![Captura de tela do fluxo "Atualizações" ao implantar um aplicativo atualizado.](../images/outlook-deploy-update-flow.png)

## <a name="see-also"></a>Confira também

- [Configurar seu suplemento do Outlook para ativação baseada em evento](autolaunch.md)
