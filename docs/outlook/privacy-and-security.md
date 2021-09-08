---
title: Privacidade, permissões e segurança de suplementos do Outlook
description: Saiba como gerenciar a privacidade, as permissões e a segurança em um suplemento do Outlook.
ms.date: 07/27/2021
localization_priority: Priority
ms.openlocfilehash: cd0c793bb8a2cfbc4cef17e0cf717c35cb68794c
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937208"
---
# <a name="privacy-permissions-and-security-for-outlook-add-ins"></a>Privacidade, permissões e segurança de suplementos do Outlook

Usuários finais, desenvolvedores e administradores podem usar os níveis de permissões em camadas do modelo de segurança para suplementos do Outlook a fim de controlar a privacidade e o desempenho.

Esse artigo descreve as possíveis permissões que os suplementos do Outlook podem solicitar e examina o modelo de segurança das seguintes perspectivas.

- **AppSource**: Integridade do suplemento

- **End-users**: Questões de privacidade e desempenho

- **Developers**: Opções de permissões e limites de uso do recurso

- **Administrators**: Privilégios para definir limites de desempenho

## <a name="permissions-model"></a>Modelo de permissões

Como a percepção dos clientes de segurança do suplemento pode afetar a sua adoção, a segurança do suplemento do Outlook conta com um modelo de permissões hierárquico. Um suplemento do Outlook divulga o nível de permissões necessárias, identificando os possíveis acessos e ações que o suplemento pode realizar em dados da caixa de correio do cliente.

A versão 1.1 do esquema do manifesto inclui quatro níveis de permissões.

**Tabela 1. Níveis de permissão do suplemento**

|**Nível de permissão**|**Valor no manifesto de suplemento do Outlook**|
|:-----|:-----|
|Restricted|Restricted|
|Leitura de item|ReadItem|
|Leitura/gravação de item|ReadWriteItem|
|Leitura/gravação de caixa de correio|ReadWriteMailbox|

Os quatro níveis de permissão são cumulativos: a permissão **leitura/gravação de caixa de correio** inclui as permissões **leitura/gravação de item**, **leitura de item** e **restrita**, **leitura/gravação de item** inclui **leitura de item** e **restrita** e a permissão **leitura de item** inclui **restrita**.

A figura a seguir mostra os quatro níveis de permissões e descreve os recursos oferecidos para o usuário final, para o desenvolvedor e para o administrador em cada nível. Para saber mais sobre essas permissões, confira [Usuários finais: questões de privacidade e desempenho](#end-users-privacy-and-performance-concerns), [Desenvolvedores: opções de permissões e limites de uso de recursos](#developers-permission-choices-and-resource-usage-limits) e [Noções básicas sobre permissões de suplementos do Outlook](understanding-outlook-add-in-permissions.md).

**Relacionando o modelo de quatro níveis de permissão com o usuário final, o desenvolvedor e o administrador**

![Modelo de permissões de 4 camadas para o esquema de aplicativos de email v1.1.](../images/add-in-permission-tiers.png)

## <a name="appsource-add-in-integrity"></a>AppSource: Integridade do suplemento

A [AppSource](https://appsource.microsoft.com) hospeda suplementos que podem ser instalados por usuários finais e administradores. A AppSource impõe as seguintes medidas para manter a integridade desses suplementos do Outlook.

- Requer que o servidor host de um suplemento sempre use o protocolo SSL para se comunicar.

- Requer que um desenvolvedor forneça uma prova de identidade, um acordo contratual e uma política de privacidade compatível para enviar suplementos.

- Suplementos de arquivos morto no modo somente leitura.

- Dá suporte a um sistema de revisão pelo usuário para os suplementos disponíveis para promover uma comunidade autovigilante.

## <a name="optional-connected-experiences"></a>Experiências conectadas opcionais

Os usuários finais e administradores de TI podem desativar as [experiências conectadas opcionais nos clientes móveis e na área de trabalho do Office](/deployoffice/privacy/optional-connected-experiences). Para suplementos do Outlook, o impacto da desabilitação da configuração das **experiências conectadas opcionais** depende do cliente, mas geralmente significa que os suplementos instalados pelo usuário e o acesso à Office Store não são permitidos. Os suplementos implantados pelo administrador de TI de uma organização por meio da [Implantação Centralizada](/microsoft-365/admin/manage/centralized-deployment-of-add-ins) ainda estarão disponíveis.

- Windows\*, Mac: o botão **Obter Suplementos** não é exibido para que os usuários não possam mais gerenciar seus suplementos ou acessar a Office Store.
- Android, iOS: a caixa de diálogo **Obter suplementos** mostra somente suplementos implantados pelo administrador.
- Navegador: a disponibilidade de suplementos e o acesso ao repositório não são afetadas, para que os usuários possam continuar a [gerenciar seus suplementos](https://support.microsoft.com/office/8f2ce816-5df4-44a5-958c-f7f9d6dabdce), incluindo aqueles implantados pelo administrador.

  > [!NOTE]
  > \* Para Windows, o suporte para essa experiência/comportamento está disponível na versão 2008 (Build 13127.20296). Para obter mais detalhes em relação à sua versão, consulte a página do histórico de atualizações do [Microsoft 365](/officeupdates/update-history-office365-proplus-by-date) e de como [encontrar a versão do cliente do Office e atualizar o canal](https://support.microsoft.com/office/932788b8-a3ce-44bf-bb09-e334518b8b19).

Para obter o comportamento geral do suplemento, confira [privacidade e segurança dos Suplementos do Office](../concepts/privacy-and-security.md#optional-connected-experiences).

## <a name="end-users-privacy-and-performance-concerns"></a>End-users: Questões de privacidade e desempenho

O modelo de segurança aborda questões de segurança, privacidade e desempenho de usuários finais das seguintes maneiras.

- Mensagens do usuário final no Outlook que são protegidas por IRM (Gerenciamento de Direitos de Informação) não interagem com os suplementos do Outlook.

  > [!IMPORTANT]
  > - Os suplementos são ativados em mensagens assinadas digitalmente no Outlook associadas a uma assinatura do Microsoft 365. No Windows, esse suporte foi introduzido com a compilação 8711.1000.
  >
  > - A partir do Outlook, build 13229.10000, no Windows, os suplementos agora podem ser ativados nos itens protegidos por IRM. Para obter mais informações sobre esse recurso na visualização, consulte [Ativação de suplementos em itens protegidos pela Gestão de Direitos de Informação (IRM)](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md#add-in-activation-on-items-protected-by-information-rights-management-irm).

- Antes de instalar um suplemento do AppSource, os usuários finais podem ver o acesso e as ações que o suplemento pode fazer em seus dados e devem confirmar explicitamente para continuar. Nenhum suplemento do Outlook é enviado automaticamente por push para um computador cliente sem validação manual pelo usuário ou administrador.

- A concessão de permissão **restrita** permite que o suplemento do Outlook tenha acesso limitado apenas ao item atual. Conceder a permissão de **item de leitura** permite que o suplemento do Outlook acesse informações de identificação pessoal, como nomes de remetentes e destinatários e endereços de email, somente no item atual.

- Um usuário final pode instalar um suplemento do Outlook somente para si mesmo. Os suplementos do Outlook que afetam uma organização são instalados por um administrador.

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

- Os desenvolvedores solicitam permissão usando o elemento [Permissions](../reference/manifest/permissions.md) no manifesto do suplemento do Outlook, atribuindo um valor **Restricted**, **ReadItem**, **ReadWriteItem** ou **ReadWriteMailbox** conforme o caso.

  > [!NOTE]
  > Observe que a permissão **ReadWriteItem** está disponível a partir do esquema de manifesto v1.1.

  Os exemplos a seguir exigem a permissão **read item**.

  ```XML
    <Permissions>ReadItem</Permissions>
  ```

- Os desenvolvedores podem solicitar a permissão **restrita** se o suplemento do Outlook for ativado em um tipo específico de itens do Outlook (compromisso ou mensagem) ou em entidades específicas extraídas (endereço número de telefone, URL) presentes no assunto ou no corpo do item. Por exemplo, a regra a seguir ativa o suplemento do Outlook se uma ou mais dessas três entidades, número de telefone, endereços postais ou URL, aparece no assunto ou no corpo da mensagem atual.

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

- Os desenvolvedores devem solicitar a permissão de **item de leitura** quando o suplemento do Outlook precisar ler as propriedades do item atual, que não sejam as entidades padrão extraídas, ou gravar propriedades personalizadas definidas pelo suplemento no item atual, mas não precisar ler ou gravar em outros itens ou criar e enviar uma mensagem na caixa de correio do usuário. Por exemplo, um desenvolvedor deve solicitar a permissão de **item de leitura** quando o suplemento do Outlook precisa procurar por uma entidade como sugestão de reunião, sugestão de tarefa, endereço de email ou nome de contato no assunto ou no corpo do item, ou usar uma expressão regular para ser ativado.

- Os desenvolvedores devem solicitar a permissão **read/write item** quando o suplemento do Outlook precisa gravar propriedades do item redigido, como nomes, endereços de email, corpo e assunto, ou precisa adicionar ou remover anexos do item.

- Os desenvolvedores solicitam a permissão **read/write mailbox** somente quando o suplemento do Outlook precisa fazer uma ou mais das seguintes ações usando o método [mailbox.makeEWSRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods).

  - Ler ou gravar em propriedades de itens na caixa de correio.
  - Criar, ler, gravar ou enviar itens na caixa de correio.
  - Criar, ler ou gravar pastas na caixa de correio.

### <a name="resource-usage-tuning"></a>Ajuste de uso do recurso

Os desenvolvedores devem estar cientes dos limites de uso do recurso para a ativação e incorporar o ajuste no seu fluxo de trabalho de desenvolvimento para reduzir a chance de ter um suplemento com mau desempenho negando serviço do host. Os desenvolvedores devem seguir as diretrizes ao criar regras de ativação conforme descrito em [Limites de ativação e API JavaScript para suplementos do Outlook](limits-for-activation-and-javascript-api-for-outlook-add-ins.md). Se um suplemento do Outlook deve ser executado em um cliente avançado do Outlook, os desenvolvedores devem verificar se o suplemento tem desempenho dentro dos limites de uso do recurso.

### <a name="other-measures-to-promote-user-security"></a>Outras medidas para promover a segurança do usuário

Os desenvolvedores devem estar atentos e planejar o seguinte.

- Desenvolvedores não podem usar controles ActiveX em suplementos porque esses não têm suporte.

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
