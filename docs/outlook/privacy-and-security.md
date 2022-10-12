---
title: Privacidade, permissões e segurança de suplementos do Outlook
description: Saiba como gerenciar a privacidade, as permissões e a segurança em um suplemento do Outlook.
ms.date: 10/07/2022
ms.localizationpriority: high
ms.openlocfilehash: 560c9bbdfcde849b66d86e9c000d78f094b3e561
ms.sourcegitcommit: a2df9538b3deb32ae3060ecb09da15f5a3d6cb8d
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/12/2022
ms.locfileid: "68541245"
---
# <a name="privacy-permissions-and-security-for-outlook-add-ins"></a>Privacidade, permissões e segurança de suplementos do Outlook

Usuários finais, desenvolvedores e administradores podem usar os níveis de permissões em camadas do modelo de segurança para suplementos do Outlook a fim de controlar a privacidade e o desempenho.

Esse artigo descreve as possíveis permissões que os suplementos do Outlook podem solicitar e examina o modelo de segurança das seguintes perspectivas.

- **AppSource**: Integridade do suplemento

- **End-users**: Questões de privacidade e desempenho

- **Developers**: Opções de permissões e limites de uso do recurso

- **Administrators**: Privilégios para definir limites de desempenho

## <a name="permissions-model"></a>Modelo de permissões

Because customers' perception of add-in security can affect add-in adoption, Outlook add-in security relies on a tiered permissions model. An Outlook add-in would disclose the level of permissions it needs, identifying the possible access and actions that the add-in can make on the customer's mailbox data.

Há quatro níveis de permissões.

[!include[Table of Outlook permissions](../includes/outlook-permission-levels-table.md)]

Os quatro níveis de permissão são cumulativos: a permissão **leitura/gravação de caixa de correio** inclui as permissões **leitura/gravação de item**, **leitura de item** e **restrita**, **leitura/gravação de item** inclui **leitura de item** e **restrita** e a permissão **leitura de item** inclui **restrita**.

A figura a seguir mostra os quatro níveis de permissões e descreve os recursos oferecidos para o usuário final, para o desenvolvedor e para o administrador em cada nível. Para saber mais sobre essas permissões, confira [Usuários finais: questões de privacidade e desempenho](#end-users-privacy-and-performance-concerns), [Desenvolvedores: opções de permissões e limites de uso de recursos](#developers-permission-choices-and-resource-usage-limits) e [Noções básicas sobre permissões de suplementos do Outlook](understanding-outlook-add-in-permissions.md).

**Relacionando o modelo de quatro níveis de permissão com o usuário final, o desenvolvedor e o administrador**

![Diagrama do modelo de permissões de quatro camadas para o esquema de aplicativos de email v1.1.](../images/add-in-permission-tiers.png)

## <a name="appsource-add-in-integrity"></a>AppSource: Integridade do suplemento

A [AppSource](https://appsource.microsoft.com) hospeda suplementos que podem ser instalados por usuários finais e administradores. A AppSource impõe as seguintes medidas para manter a integridade desses suplementos do Outlook.

- Requer que o servidor host de um suplemento sempre use o protocolo SSL para se comunicar.

- Requer que um desenvolvedor forneça uma prova de identidade, um acordo contratual e uma política de privacidade compatível para enviar suplementos.

- Suplementos de arquivos morto no modo somente leitura.

- Dá suporte a um sistema de revisão pelo usuário para os suplementos disponíveis para promover uma comunidade autovigilante.

## <a name="optional-connected-experiences"></a>Experiências conectadas opcionais

Os usuários finais e administradores de TI podem desativar as [experiências conectadas opcionais nos clientes móveis e na área de trabalho do Office](/deployoffice/privacy/optional-connected-experiences). Para suplementos do Outlook, o impacto de desabilitar a configuração  de experiências conectadas opcionais depende do cliente, mas geralmente significa que suplementos instalados pelo usuário e acesso à Office Store não são permitidos. Os suplementos implantados pelo administrador de TI de uma organização por meio da [Implantação Centralizada](/microsoft-365/admin/manage/centralized-deployment-of-add-ins) ainda estarão disponíveis.

- Windows\*, Mac: o **botão Obter Suplementos** não é exibido para que os usuários não possam mais gerenciar seus suplementos ou acessar a Office Store.
- Android, iOS: a caixa de diálogo **Obter suplementos** mostra somente suplementos implantados pelo administrador.
- Navegador: a disponibilidade de suplementos e o acesso ao repositório não são afetadas, para que os usuários possam continuar a [gerenciar seus suplementos](https://support.microsoft.com/office/8f2ce816-5df4-44a5-958c-f7f9d6dabdce), incluindo aqueles implantados pelo administrador.

  > [!NOTE]
  > \* Para Windows, o suporte para essa experiência/comportamento está disponível na versão 2008 (Build 13127.20296). Para obter mais detalhes em relação à sua versão, consulte a página do histórico de atualizações do [Microsoft 365](/officeupdates/update-history-office365-proplus-by-date) e de como [encontrar a versão do cliente do Office e atualizar o canal](https://support.microsoft.com/office/932788b8-a3ce-44bf-bb09-e334518b8b19).

Para obter o comportamento geral do suplemento, confira [privacidade e segurança dos Suplementos do Office](../concepts/privacy-and-security.md#optional-connected-experiences).

## <a name="end-users-privacy-and-performance-concerns"></a>End-users: Questões de privacidade e desempenho

O modelo de segurança aborda questões de segurança, privacidade e desempenho de usuários finais das seguintes maneiras.

- As mensagens do usuário final protegidas pelo IRM (Gerenciamento de Direitos de Informação) do Outlook não interagem com suplementos do Outlook em clientes não Windows.

[!INCLUDE [outlook-irm-add-in-activation](../includes/outlook-irm-add-in-activation.md)]

- Before installing an add-in from AppSource, end users can see the access and actions that the add-in can make on their data and must explicitly confirm to proceed. No Outlook add-in is automatically pushed onto a client computer without manual validation by the user or administrator.

- A concessão da permissão **restricted** permite que o suplemento do Outlook tenha acesso limitado apenas ao item atual. A concessão da permissão **de item** de leitura permite que o suplemento do Outlook acesse informações de identificação pessoal, como nomes de remetente e destinatário e endereços de email, somente no item atual.

- An end user can install an Outlook add-in for only himself or herself. Outlook add-ins that affect an organization are installed by an administrator.

- Os usuários finais podem instalar suplementos do Outlook que permitem cenários dependentes do contexto, o que é atraente para os usuários e reduz os riscos de segurança.

- Arquivos de manifesto de suplementos do Outlook instalados são protegidos na conta de email do usuário.

- Dados comunicados com servidores que hospedam os Suplementos do Office são sempre criptografados de acordo com o protocolo SSL (Secure Socket Layer).

- Aplicável apenas aos clientes avançados do Outlook: Os clientes avançados do Outlook monitoram o desempenho de suplementos do Outlook instalados, exercem controle de governança e desabilitam os suplementos do Outlook que excedem os limites nas seguintes áreas.

  - Tempo de resposta para ativação

  - Número de falhas na ativação ou reativação

  - Uso da memória

  - Uso da CPU  

  Governance deters denial-of-service attacks and maintains add-in performance at a reasonable level. The Business Bar alerts end users about Outlook add-ins that the Outlook rich client has disabled based on such governance control.

- A qualquer hora, os usuários finais podem verificar as permissões solicitadas pelos suplementos do Outlook instalados e desabilitar ou habilitar subsequentemente qualquer suplemento do Outlook no Centro de Administração do Exchange.

## <a name="developers-permission-choices-and-resource-usage-limits"></a>Desenvolvedores: opções de permissões e limites de uso do recurso.

O modelo de segurança fornece aos desenvolvedores níveis granulares de permissão à sua escolha e diretrizes de desempenho rígidas a observar.

### <a name="tiered-permissions-increases-transparency"></a>Permissões hierárquicas aumentam a transparência

Os desenvolvedores devem seguir o modelo de permissões hierárquico para dar transparência e aliviar as preocupações dos usuários em relação ao que os suplementos podem fazer por seus dados e caixa de correio, promovendo indiretamente a adoção do suplemento.

- Os desenvolvedores solicitam um nível adequado de permissão para um suplemento do Outlook, com base em como o suplemento do Outlook deve ser ativado e na sua necessidade de ler ou gravar determinadas propriedades de um item, ou de criar e enviar um item.

- Conforme mencionado acima, os desenvolvedores solicitam permissão no manifesto.

  O exemplo a seguir solicita **a permissão de leitura de item** no manifesto XML.

  ```XML
  <Permissions>ReadItem</Permissions>
  ```

  O exemplo a seguir solicita **a permissão de leitura de item** no manifesto do Teams (versão prévia).

```json
"authorization": {
  "permissions": {
    "resourceSpecific": [
      ...
      {
        "name": "MailboxItem.Read.User",
        "type": "Delegated"
      },
    ]
  }
},
```

- Os desenvolvedores podem  solicitar a permissão restrita se o suplemento do Outlook for ativado em um tipo específico de item do Outlook (compromisso ou mensagem) ou em entidades extraídas específicas (número de telefone, endereço, URL) presentes no assunto ou no corpo do item. Por exemplo, a regra a seguir ativa o suplemento do Outlook se uma ou mais dessas três entidades, número de telefone, endereços postais ou URL, aparece no assunto ou no corpo da mensagem atual.

> [!NOTE]
> As regras de ativação, como visto neste exemplo, não têm suporte em suplementos que usam o manifesto do [Teams para Suplementos do Office (versão prévia)](../develop/json-manifest-overview.md).

  ```XML
    <Permissions>Restricted</Permissions>
        <Rule xsi:type="RuleCollection" Mode="And">
        <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
        <Rule xsi:type="RuleCollection" Mode="Or">
            <Rule xsi:type="ItemHasKnownEntity" EntityType="PhoneNumber" />
            <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
            <Rule xsi:type="ItemHasKnownEntity" EntityType="Url" />
        </Rule>
    </Rule>
  ```

- Os desenvolvedores devem solicitar a permissão de **item** de leitura se o suplemento do Outlook precisar ler propriedades do item atual que não sejam as entidades extraídas padrão ou gravar propriedades personalizadas definidas pelo suplemento no item atual, mas não exigir leitura ou gravação em outros itens ou criar ou enviar uma mensagem na caixa de correio do usuário. Por exemplo, um desenvolvedor deve solicitar a permissão **read item** quando o suplemento do Outlook precisa procurar por uma entidade como sugestão de reunião, sugestão de tarefa, endereço de email ou nome de contato no assunto ou no corpo do item, ou usar uma expressão regular para ser ativado.

- Os desenvolvedores devem solicitar a permissão **read/write item** quando o suplemento do Outlook precisa gravar propriedades do item redigido, como nomes, endereços de email, corpo e assunto, ou precisa adicionar ou remover anexos do item.

- Os desenvolvedores solicitam a permissão **read/write mailbox** somente quando o suplemento do Outlook precisa fazer uma ou mais das seguintes ações usando o método [mailbox.makeEWSRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods).

  - Ler ou gravar em propriedades de itens na caixa de correio.
  - Criar, ler, gravar ou enviar itens na caixa de correio.
  - Criar, ler ou gravar pastas na caixa de correio.

### <a name="resource-usage-tuning"></a>Ajuste de uso do recurso

Developers should be aware of resource usage limits for activation, incorporate performance tuning in their development workflow, so as to reduce the chance of a poorly performing add-in denying service of the host. Developers should follow the guidelines in designing activation rules as described in [Limits for activation and JavaScript API for Outlook add-ins](limits-for-activation-and-javascript-api-for-outlook-add-ins.md). If an Outlook add-in is intended to run on an Outlook rich client, then developers should verify that the add-in performs within the resource usage limits.

### <a name="other-measures-to-promote-user-security"></a>Outras medidas para promover a segurança do usuário

Os desenvolvedores devem estar atentos e planejar o seguinte.

- Os desenvolvedores não podem usar controles ActiveX em suplementos porque não têm suporte.

- Os desenvolvedores devem fazer o seguinte ao enviar um suplemento do Outlook à AppSource.

  - Criar um certificado SSL EV (validação estendida) como prova de identidade.

  - Hospedar o suplemento que estão enviando em um servidor Web que dê suporte a SSL.

  - Criar uma política de privacidade compatível.

  - Estar preparados para assinar um acordo contratual ao enviar o suplemento.

## <a name="administrators-privileges"></a>Administradores: privilégios

O modelo de segurança fornece os seguintes direitos e responsabilidades aos administradores.

- Podem impedir que os usuários finais instalem suplementos do Outlook, incluindo suplementos da AppSource.

- Podem habilitar ou desabilitar qualquer suplemento do Outlook no Centro de Administração do Exchange.

- Aplicável apenas ao Outlook no Windows: pode substituir as configurações de limite de desempenho por configurações de registro de GPO.

## <a name="see-also"></a>Confira também

- [Privacidade e segurança para Suplementos do Office](../concepts/privacy-and-security.md)
- [Controles de privacidade para Microsoft 365 Apps](/deployoffice/privacy/overview-privacy-controls)
- [APIs de suplemento do Outlook](apis.md)
- [Limites para ativação e API JavaScript para suplementos do Outlook](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
