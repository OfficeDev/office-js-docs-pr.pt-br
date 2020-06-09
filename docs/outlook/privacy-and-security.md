---
title: Privacidade, permissões e segurança de suplementos do Outlook
description: Saiba como gerenciar a privacidade, as permissões e a segurança em um suplemento do Outlook.
ms.date: 10/31/2019
localization_priority: Priority
ms.openlocfilehash: d233eb3ac6980af24e6ba9d951834532ea79dc06
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44605329"
---
# <a name="privacy-permissions-and-security-for-outlook-add-ins"></a>Privacidade, permissões e segurança de suplementos do Outlook

Usuários finais, desenvolvedores e administradores podem usar os níveis de permissões em camadas do modelo de segurança para suplementos do Outlook a fim de controlar a privacidade e o desempenho.

Este artigo descreve as possíveis permissões que os suplementos do Outlook podem solicitar e examina o modelo de segurança das seguintes perspectivas:

- **AppSource**: integridade do suplemento
    
- **Usuários finais**: questões de privacidade e desempenho
    
- **Desenvolvedores**: opções de permissões e limites de uso do recurso
    
- **Administradores**: privilégios para definir limites de desempenho
    

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

![Modelo de permissões de quatro camadas para o esquema de aplicativos de correio v1.1](../images/add-in-permission-tiers.png)


## <a name="appsource-add-in-integrity"></a>AppSource: integridade do suplemento

A [AppSource](https://appsource.microsoft.com) hospeda suplementos que podem ser instalados por usuários finais e administradores. A AppSource impõe as seguintes medidas para manter a integridade desses suplementos do Outlook:

- Requer que o servidor host de um suplemento sempre use o protocolo SSL para se comunicar.
    
- Requer que um desenvolvedor forneça uma prova de identidade, um acordo contratual e uma política de privacidade compatível para enviar suplementos. 
    
- Suplementos de arquivos morto no modo somente leitura.
    
- Dá suporte a um sistema de revisão pelo usuário para os suplementos disponíveis para promover uma comunidade autovigilante.
    

## <a name="end-users-privacy-and-performance-concerns"></a>Usuários finais: questões de privacidade e desempenho.

O modelo de segurança aborda questões de segurança, privacidade e desempenho de usuários finais das seguintes maneiras:

- Mensagens do usuário final no Outlook que são protegidas por IRM (Gerenciamento de Direitos de Informação) não interagem com os suplementos do Outlook.
    
- Antes de instalar um suplemento da AppSource, os usuários finais podem ver o acesso e as ações que o suplemento pode realizar em seus dados e deve confirmá-los explicitamente para prosseguir. Nenhum suplemento do Outlook é automaticamente enviado por push a um computador cliente sem validação manual pelo usuário ou administrador.
    
- A concessão da permissão **restricted** permite que o suplemento do Outlook tenha acesso limitado apenas ao item atual. A concessão da permissão **read item** permite que o suplemento do Outlook acesse informações de identificação pessoal, como remetente e nomes dos destinatários e endereços de email, apenas no item atual.
    
- Um usuário final pode instalar um suplemento do Outlook somente para si mesmo. Os suplementos do Outlook que afetam uma organização são instalados por um administrador.
    
- Os usuários finais podem instalar suplementos do Outlook que permitem cenários dependentes do contexto, o que é atraente para os usuários e reduz os riscos de segurança.
    
- Arquivos de manifesto de suplementos do Outlook instalados são protegidos na conta de email do usuário.
    
- Dados comunicados com servidores que hospedam os Suplementos do Office são sempre criptografados de acordo com o protocolo SSL (Secure Socket Layer).
    
- Aplicável apenas aos clientes avançados do Outlook: Os clientes avançados do Outlook monitoram o desempenho de suplementos do Outlook instalados, exercem controle de governança e desabilitam os suplementos do Outlook que excedem os limites nas seguintes áreas:
    
  - Tempo de resposta para ativação
    
  - Número de falhas na ativação ou reativação
    
  - Uso da memória
    
  - Uso da CPU  

  Governance deters denial-of-service attacks and maintains add-in performance at a reasonable level. The Business Bar alerts end users about Outlook add-ins that the Outlook rich client has disabled based on such governance control.

- A qualquer hora, os usuários finais podem verificar as permissões solicitadas pelos suplementos do Outlook instalados e desabilitar ou habilitar subsequentemente qualquer suplemento do Outlook no Centro de Administração do Exchange.


## <a name="developers-permission-choices-and-resource-usage-limits"></a>Desenvolvedores: opções de permissões e limites de uso do recurso.

O modelo de segurança fornece aos desenvolvedores níveis granulares de permissão à sua escolha e diretrizes de desempenho rígidas a observar.

### <a name="tiered-permissions-increases-transparency"></a>Permissões hierárquicas aumentam a transparência

Os desenvolvedores devem seguir o modelo de permissões hierárquico para dar transparência e aliviar as preocupações dos usuários em relação ao que os suplementos podem fazer por seus dados e caixa de correio, promovendo indiretamente a adoção do suplemento:

- Os desenvolvedores solicitam um nível adequado de permissão para um suplemento do Outlook, com base em como o suplemento do Outlook deve ser ativado e na sua necessidade de ler ou gravar determinadas propriedades de um item, ou de criar e enviar um item.

- Os desenvolvedores solicitam permissão usando o elemento [Permissions](../reference/manifest/permissions.md) no manifesto do suplemento do Outlook, atribuindo um valor **Restricted**, **ReadItem**, **ReadWriteItem** ou **ReadWriteMailbox** conforme o caso.

  > [!NOTE]
  > Observe que a permissão **ReadWriteItem** está disponível a partir do esquema de manifesto v1.1.

  Os exemplos a seguir exigem a permissão **read item**.

  ```XML
    <Permissions>ReadItem</Permissions>
  ```

- Os desenvolvedores podem solicitar a permissão **restricted** se o suplemento do Outlook for ativado em um tipo específico de itens do Outlook (compromisso ou mensagem) ou em entidades específicas extraídas (endereço número de telefone, URL) presentes no assunto ou no corpo do item. Por exemplo, a regra a seguir ativa o suplemento do Outlook se uma ou mais dessas três entidades, número de telefone, endereços postais ou URL, aparece no assunto ou no corpo da mensagem atual.
    
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

- Os desenvolvedores devem solicitar a permissão **read item** quando o suplemento do Outlook precisa ler as propriedades do item atual, que não sejam as entidades padrão extraídas, ou gravar propriedades personalizadas definidas pelo suplemento no item atual, mas não precisa ler ou gravar em outros itens ou criar e enviar uma mensagem na caixa de correio do usuário. Por exemplo, um desenvolvedor deve solicitar a permissão **read item** quando o suplemento do Outlook precisa procurar por uma entidade como sugestão de reunião, sugestão de tarefa, endereço de email ou nome de contato no assunto ou no corpo do item, ou usar uma expressão regular para ser ativado.

- Os desenvolvedores devem solicitar a permissão **read/write item** quando o suplemento do Outlook precisa gravar propriedades do item redigido, como nomes, endereços de email, corpo e assunto, ou precisa adicionar ou remover anexos do item.

- Os desenvolvedores solicitam a permissão **read/write mailbox** somente quando o suplemento do Outlook precisa fazer uma ou mais das seguintes ações usando o método [mailbox.makeEWSRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods):

  - Ler ou gravar em propriedades de itens na caixa de correio.
  - Criar, ler, gravar ou enviar itens na caixa de correio.
  - Criar, ler ou gravar pastas na caixa de correio.


### <a name="resource-usage-tuning"></a>Ajuste de uso do recurso

Os desenvolvedores devem estar cientes dos limites de uso do recurso para a ativação e incorporar o ajuste no seu fluxo de trabalho de desenvolvimento para reduzir a chance de ter um suplemento com mau desempenho negando serviço do host. Os desenvolvedores devem seguir as diretrizes ao criar regras de ativação conforme descrito em [Limites de ativação e API JavaScript para suplementos do Outlook](limits-for-activation-and-javascript-api-for-outlook-add-ins.md). Se um suplemento do Outlook deve ser executado em um cliente avançado do Outlook, os desenvolvedores devem verificar se o suplemento tem desempenho dentro dos limites de uso do recurso.


### <a name="other-measures-to-promote-user-security"></a>Outras medidas para promover a segurança do usuário

Os desenvolvedores devem estar atentos e planejar o seguinte:

- Desenvolvedores não podem usar controles ActiveX em suplementos porque esses não têm suporte.
    
- Os desenvolvedores devem fazer o seguinte ao enviar um suplemento do Outlook à AppSource:
    
  - Criar um certificado SSL EV (validação estendida) como prova de identidade.
    
  - Hospedar o suplemento que estão enviando em um servidor Web que dê suporte a SSL.
    
  - Criar uma política de privacidade compatível.
    
  - Estar preparados para assinar um acordo contratual ao enviar o suplemento.
    

## <a name="administrators-privileges"></a>Administradores: privilégios

O modelo de segurança fornece os seguintes direitos e responsabilidades aos administradores:

- Podem impedir que os usuários finais instalem suplementos do Outlook, incluindo suplementos da AppSource.
    
- Podem habilitar ou desabilitar qualquer suplemento do Outlook no Centro de Administração do Exchange.
    
- Aplicável apenas ao Outlook no Windows: pode substituir as configurações de limite de desempenho por configurações de registro de GPO.
    


## <a name="see-also"></a>Confira também

- [Privacidade e segurança para Suplementos do Office](../develop/privacy-and-security.md)    
- [APIs de suplemento do Outlook](apis.md)    
- [Limites para ativação e API JavaScript para suplementos do Outlook](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
    
