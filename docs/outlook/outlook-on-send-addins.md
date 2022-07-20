---
title: Recurso Ao enviar para suplementos do Outlook
description: Fornece uma maneira de manipular um item ou impedir que usuários realizem determinadas ações e permite que um suplemento defina determinadas propriedades ao enviar.
ms.date: 07/14/2022
ms.localizationpriority: medium
ms.openlocfilehash: 5a5b9d964c48496658157b4a8506bf283419fbb2
ms.sourcegitcommit: df7964b6509ee6a807d754fbe895d160bc52c2d3
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/20/2022
ms.locfileid: "66889601"
---
# <a name="on-send-feature-for-outlook-add-ins"></a>Recurso Ao enviar para suplementos do Outlook

O recurso Ao enviar para suplementos do Outlook fornece uma maneira de manipular uma mensagem ou item de reunião, ou impede que usuários realizem determinadas ações e permite que um suplemento defina determinadas propriedades ao enviar. Por exemplo, você pode usar o recurso Ao enviar para:

- Impedir que um usuário envie informações confidenciais ou deixe a linha de assunto em branco.  
- Adicionar um destinatário específico à linha CC em mensagens ou à linha destinatários opcionais em reuniões.

O recurso ao enviar é acionado pelo tipo de evento `ItemSend` e é sem interface de usuário.

Para obter informações sobre limitações relacionadas ao recurso Ao enviar, consulte as [Limitações](#limitations) posteriormente neste artigo.

## <a name="supported-clients-and-platforms"></a>Clientes e plataformas com suporte

A tabela a seguir mostra combinações de cliente-servidor com suporte para o recurso ao enviar, incluindo a atualização cumulativa mínima necessária, quando aplicável. Não há suporte para combinações excluídas.

| Client | Exchange Online | Exchange 2016 local<br>(Atualização Cumulativa 6 ou posterior) | Exchange 2019 local<br>(Atualização Cumulativa 1 ou posterior) |
|---|:---:|:---:|:---:|
|Windows:<br>versão 1910 (build 12130.20272) ou posterior|Sim|Sim|Sim|
|Mac:<br>build 16.47 ou posterior|Sim|Sim|Sim|
|Navegador da Web:<br>interface do usuário moderna do Outlook|Sim|Não aplicável|Não aplicável|
|Navegador da Web:<br>interface do usuário clássica do Outlook|Não aplicável|Sim|Sim|

> [!NOTE]
> O recurso ao enviar foi lançado oficialmente no conjunto de requisitos 1.8 (consulte o servidor [atual e o suporte ao cliente](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients) para obter detalhes). No entanto, observe que a matriz de suporte do recurso é um superconjunto do conjunto de requisitos.

> [!IMPORTANT]
> Os suplementos que usam o recurso ao enviar não são permitidos no [AppSource](https://appsource.microsoft.com).

## <a name="how-does-the-on-send-feature-work"></a>Como o recurso Ao enviar funciona?

Você pode usar o recurso Ao enviar para criar um suplemento do Outlook que integre o evento síncrono `ItemSend`. Este evento detecta que o usuário está pressionando o botão **Enviar** (ou o botão **Enviar Atualização** para reuniões existentes) e pode ser usado para impedir que um item seja enviado se houver falha na validação. Por exemplo, quando um usuário dispara um evento de envio de mensagem, um suplemento do Outlook que usa o recurso Ao enviar pode:

- Leia e valide o conteúdo da mensagem de email.
- Verifique se a mensagem inclui uma linha de assunto.
- Defina um destinatário predeterminado.

A validação é feita no lado do cliente no Outlook quando o evento de envio é disparado e o suplemento tem até 5 minutos antes de expirar. Se a validação falhar, o envio do item será bloqueado e uma mensagem de erro será exibida em uma barra de informações que solicita que o usuário execute uma ação.

> [!NOTE]
> No Outlook na Web, quando o recurso ao enviar é disparado em uma mensagem que está sendo composta na guia do navegador Outlook, o item é exibido em sua própria janela ou guia do navegador para concluir a validação e outros processamentos.

A captura de tela a seguir mostra uma barra de informações que notifica que o remetente adicione um assunto.

![Uma mensagem de erro solicitando que o usuário insira uma linha de assunto ausente.](../images/block-on-send-subject-cc-inforbar.png)

A captura de tela a seguir mostra uma barra de informações que notifica que o remetente de que foram encontradas palavras bloqueadas.

![Uma mensagem de erro informando ao usuário que foram encontradas palavras bloqueadas.](../images/block-on-send-body.png)

## <a name="limitations"></a>Limitações

Atualmente, o recurso Ao enviar tem as seguintes limitações.

- **Recurso Append-on-send** &ndash; Se você chamar [item.body.AppendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-1.9&preserve-view=true#outlook-office-body-appendonsendasync-member(1)) no manipulador ao enviar, um erro será retornado.
- **AppSource** &ndash; Você não pode publicar suplementos do Outlook que usem o recurso Ao enviar no [AppSource](https://appsource.microsoft.com), pois eles falharão na validação do AppSource. Os suplementos que usam o recurso Ao enviar devem ser implantados pelos administradores.
  
  > [!IMPORTANT]
  > Ao executar `npm run validate` para [validar](../testing/troubleshoot-manifest.md) o manifesto do suplemento, você receberá o erro "O suplemento caixa de correio que contém o evento ItemSend é inválido. O manifesto do suplemento caixa de correio contém o evento ItemSend em VersionOverrides, o que não é permitido." Essa mensagem é exibida porque os suplementos que usam o evento, que é necessário para esta versão do recurso ao enviar, não podem ser publicados `ItemSend` no AppSource. Você ainda poderá realizar o sideload e executar o suplemento, desde que nenhum outro erro de validação seja encontrado.

- **Manifesto** &ndash; Somente um evento `ItemSend` tem suporte por suplemento. Se você tiver dois ou mais eventos `ItemSend` em um manifesto, haverá falha na validação.
- **Desempenho**&ndash; Várias idas e voltas ao servidor Web que hospeda o suplemento podem afetar o desempenho do suplemento. Considere os efeitos sobre o desempenho quando você cria suplemento que exigem várias mensagens ou operações baseadas em reuniões.
- **Enviar mais tarde** (somente Mac) &ndash; Se houver suplementos Ao enviar, o recurso **Enviar mais tarde** ficará indisponível.

Além disso, não é recomendável `item.close()` que você chame o manipulador de eventos ao enviar, pois o fechamento do item deve ocorrer automaticamente após a conclusão do evento.

### <a name="mailbox-typemode-limitations"></a>Limitações de tipo/modo de caixa de correio

A funcionalidade Ao enviar é compatível apenas com caixas de correio de usuários no Outlook na Web, Windows e Mac. Além das situações em que os suplementos não são ativados conforme descrito nos itens de Caixa de Correio disponíveis para [a seção suplementos](outlook-add-ins-overview.md#mailbox-items-available-to-add-ins) da página de visão geral de suplementos do Outlook, a funcionalidade não tem suporte no momento para o modo offline em que esse modo está disponível.

Nos casos em que os suplementos do Outlook não são ativados, o suplemento ao enviar não será executado e a mensagem será enviada.

No entanto, se o recurso ao enviar estiver habilitado e disponível, mas o cenário de caixa de correio não tiver suporte, o Outlook não permitirá o envio.

## <a name="multiple-on-send-add-ins"></a>Vários suplementos Ao enviar

Se vários suplementos Ao enviar estiverem instalados, os suplementos serão executados na ordem em que são recebidos das APIs `getAppManifestCall` ou `getExtensibilityContext`. Se o primeiro suplemento permitir envio, o segundo suplemento poderá alterar algo que faria o primeiro bloquear o envio. No entanto, o primeiro suplemento não será executado novamente se todos os suplementos instalados tiverem permissão de envio.

Por exemplo, o Suplemento1 e o Suplemento2 usam o recurso Ao enviar. O Suplemento1 é instalado primeiro e o Suplemento2 é instalado depois. O Suplemento1 verifica se a palavra Fabrikam aparece na mensagem como uma condição para o suplemento permitir o envio.  No entanto, o Suplemento2 remove as ocorrências da palavra Fabrikam. A mensagem será enviada com todas as instâncias de Fabrikam removidas (devido à ordem de instalação do Suplemento1 e do Suplemento2).

## <a name="deploy-outlook-add-ins-that-use-on-send"></a>Implantar suplementos do Outlook que usam Ao enviar

Recomendamos que os administradores implantem suplementos do Outlook que usam o recurso Ao enviar. Os administradores precisam garantir que o suplemento Ao enviar:

- Esteja sempre presente a qualquer momento que um item de redigir é aberto (para email: novo, responder ou encaminhar).
- Não pode ser fechado ou desabilitado pelo usuário.

## <a name="install-outlook-add-ins-that-use-on-send"></a>Instalar suplementos do Outlook que usam Ao enviar

O recurso Ao enviar no Outlook exige que os suplementos sejam configurados para os tipos de eventos de envio. Selecione a plataforma que você deseja configurar.

### <a name="web-browser---classic-outlook"></a>[Navegador da Web – Outlook clássico](#tab/classic)

Os suplementos para Outlook na Web (clássico) que usam o recurso ao enviar serão executados para usuários que recebem uma política de caixa de correio do Outlook na Web que tem o sinalizador *OnSendAddinsEnabled* definido como `true`.

Para instalar um novo suplemento, execute os seguintes cmdlets do PowerShell do Exchange Online.

```powershell
$Data=Get-Content -Path '.\Contoso Message Body Checker.xml' -Encoding Byte –ReadCount 0
```

```powershell
New-App -OrganizationApp -FileData $Data -DefaultStateForUser Enabled
```

> [!NOTE]
> Para saber como usar o PowerShell para se conectar ao Exchange Online, confira [Conectar ao Exchange Online PowerShell](/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell).

#### <a name="enable-the-on-send-feature"></a>Habilitar o recurso Ao enviar

Por padrão, a funcionalidade Ao enviar está desabilitada. Os administradores podem habilitar a funcionalidade Ao enviar executando os cmdlets do PowerShell do Exchange Online.

Para habilitar suplementos Ao enviar para todos os usuários:

1. Criar uma nova política de caixa de correio do Outlook na Web.

   ```powershell
    New-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy
   ```

    > [!NOTE]
    > Os administradores podem usar uma diretiva existente, mas a funcionalidade Ao enviar tem suporte apenas para certos tipos de caixa de correio. As caixas de correio sem suporte serão impedidas de enviar por padrão no Outlook na Web.

1. Habilitar o recurso Ao enviar.

   ```powershell
    Get-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$true
   ```

1. Atribua a política aos usuários.

   ```powershell
    Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'}|Set-CASMailbox -OwaMailboxPolicy OWAOnSendAddinAllUserPolicy
   ```

#### <a name="enable-the-on-send-feature-for-a-group-of-users"></a>Habilitar o recurso Ao enviar para um grupo de usuários

Para habilitar o recurso Ao enviar para um grupo específico de usuários, as etapas são as seguintes.  Neste exemplo, um administrador deseja habilitar apenas o recurso de suplemento Ao enviar do Outlook na Web em um ambiente para usuários do Finance (em que os usuários do Finance estão no Departamento Financeiro).

1. Crie uma nova política de caixa de correio do Outlook na Web para o grupo.

   ```powershell
    New-OWAMailboxPolicy FinanceOWAPolicy
   ```

   > [!NOTE]
   > Os administradores podem usar uma política existente, mas a funcionalidade Ao enviar é compatível apenas com certos tipos de caixa de correio (consulte [Limitações de tipo de caixa de correio](#multiple-on-send-add-ins) anteriormente neste artigo para obter mais informações). As caixas de correio sem suporte serão impedidas de enviar por padrão no Outlook na Web.

1. Habilitar o recurso Ao enviar.

   ```powershell
    Get-OWAMailboxPolicy FinanceOWAPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$true
   ```

1. Atribua a política aos usuários.

   ```powershell
    $targetUsers = Get-Group 'Finance'|select -ExpandProperty members
    $targetUsers | Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'}|Set-CASMailbox -OwaMailboxPolicy FinanceOWAPolicy
   ```

> [!NOTE]
> Espere até 60 minutos para a política entrar em vigor ou reinicie os Serviços de Informações da Internet (IIS). Quando a política entrar em vigor, o recurso Ao enviar será habilitado para o grupo.

#### <a name="disable-the-on-send-feature"></a>Desabilitar o recurso Ao enviar

Para desabilitar o recurso Ao enviar de um usuário ou atribuir uma política de caixa de correio do Outlook na Web que não tenha o sinalizador habilitado, execute os seguintes cmdlets. Neste exemplo, a política de caixa de correio é *ContosoCorpOWAPolicy*.

```powershell
Get-CASMailbox joe@contoso.com | Set-CASMailbox –OWAMailboxPolicy "ContosoCorpOWAPolicy"
```

> [!NOTE]
> Para saber mais sobre como usar o cmdlet **Set-OwaMailboxPolicy** para configurar as políticas de caixa de correio da Web existentes do Outlook, confira [Set-OwaMailboxPolicy](/powershell/module/exchange/client-access/Set-OwaMailboxPolicy).

Para desabilitar o recurso Ao enviar para todos os usuários que tenham uma política específica de caixa de correio do Outlook na Web atribuída, execute os seguintes cmdlets.

```powershell
Get-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$false
```

### <a name="web-browser---modern-outlook"></a>[Navegador da Web – Outlook moderno](#tab/modern)

Os suplementos para Outlook na Web (modernos) que usam o recurso Ao enviar devem ser executados para qualquer usuário que os tenha instalado. No entanto, se os usuários precisarem executar suplementos ao enviar para atender aos padrões de conformidade, a política de caixa de correio deverá ter o sinalizador *OnSendAddinsEnabled* `true` definido para que a edição do item não seja permitida enquanto os suplementos estiverem sendo processadas no envio.

Para instalar um novo suplemento, execute os seguintes cmdlets do PowerShell do Exchange Online.

```powershell
$Data=Get-Content -Path '.\Contoso Message Body Checker.xml' -Encoding Byte –ReadCount 0
```

```powershell
New-App -OrganizationApp -FileData $Data -DefaultStateForUser Enabled
```

> [!NOTE]
> Para saber como usar o PowerShell para se conectar ao Exchange Online, confira [Conectar ao Exchange Online PowerShell](/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell).

#### <a name="enable-the-on-send-flag"></a>Habilitar o sinalizador ao enviar

Os administradores podem impor a conformidade ao enviar executando Exchange Online cmdlets do PowerShell.

Para todos os usuários, não permitir a edição durante o processamento de suplementos ao enviar:

1. Criar uma nova política de caixa de correio do Outlook na Web.

   ```powershell
    New-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy
   ```

    > [!NOTE]
    > Os administradores podem usar uma diretiva existente, mas a funcionalidade Ao enviar tem suporte apenas para certos tipos de caixa de correio. As caixas de correio sem suporte serão impedidas de enviar por padrão no Outlook na Web.

1. Impor a conformidade ao enviar.

   ```powershell
    Get-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$true
   ```

1. Atribua a política aos usuários.

   ```powershell
    Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'}|Set-CASMailbox -OwaMailboxPolicy OWAOnSendAddinAllUserPolicy
   ```

#### <a name="turn-on-the-on-send-flag-for-a-group-of-users"></a>Ativar o sinalizador ao enviar para um grupo de usuários

Para impor a conformidade ao enviar para um grupo específico de usuários, as etapas são as seguintes. Neste exemplo, um administrador apenas deseja habilitar uma política de suplemento Ao enviar do Outlook na Web em um ambiente para usuários do Finanças (em que os usuários do Finanças estão no Departamento Financeiro).

1. Crie uma nova política de caixa de correio do Outlook na Web para o grupo.

   ```powershell
    New-OWAMailboxPolicy FinanceOWAPolicy
   ```

   > [!NOTE]
   > Os administradores podem usar uma política existente, mas a funcionalidade Ao enviar é compatível apenas com certos tipos de caixa de correio (consulte [Limitações de tipo de caixa de correio](#multiple-on-send-add-ins) anteriormente neste artigo para obter mais informações). As caixas de correio sem suporte serão impedidas de enviar por padrão no Outlook na Web.

1. Impor a conformidade ao enviar.

   ```powershell
    Get-OWAMailboxPolicy FinanceOWAPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$true
   ```

1. Atribua a política aos usuários.

   ```powershell
    $targetUsers = Get-Group 'Finance'|select -ExpandProperty members
    $targetUsers | Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'}|Set-CASMailbox -OwaMailboxPolicy FinanceOWAPolicy
   ```

> [!NOTE]
> Espere até 60 minutos para a política entrar em vigor ou reinicie os Serviços de Informações da Internet (IIS). Quando a política entrar em vigor, a conformidade ao enviar será imposta para o grupo.

#### <a name="turn-off-the-on-send-flag"></a>Desativar o sinalizador ao enviar

Para desativar a imposição de conformidade ao enviar para um usuário, atribua uma política de caixa de correio Outlook na Web que não tenha o sinalizador habilitado executando os cmdlets a seguir. Neste exemplo, a política de caixa de correio é *ContosoCorpOWAPolicy*.

```powershell
Get-CASMailbox joe@contoso.com | Set-CASMailbox –OWAMailboxPolicy "ContosoCorpOWAPolicy"
```

> [!NOTE]
> Para saber mais sobre como usar o cmdlet **Set-OwaMailboxPolicy** para configurar as políticas de caixa de correio da Web existentes do Outlook, confira [Set-OwaMailboxPolicy](/powershell/module/exchange/client-access/Set-OwaMailboxPolicy).

Para desativar a imposição de conformidade ao enviar para todos os usuários que têm uma política Outlook na Web caixa de correio específica atribuída, execute os cmdlets a seguir.

```powershell
Get-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$false
```

### <a name="windows"></a>[Windows](#tab/windows)

Os suplementos para Outlook no Windows que usam o recurso Ao enviar devem ser executados para qualquer usuário que os tenha instalado. No entanto, se os usuários precisarem executar o suplemento para atender aos padrões de conformidade, a política de grupo Bloquear envio quando os **suplementos da Web** não puderem ser carregados deverá ser definida  como Habilitada em cada computador aplicável.

Para definir políticas de caixa de correio, os administradores podem baixar a ferramenta Modelos [Administrativos](https://www.microsoft.com/download/details.aspx?id=49030) e acessar os modelos administrativos mais recentes executando o Editor local Política de Grupo, **gpedit.msc**.

> [!NOTE]
> Em versões mais antigas da ferramenta Modelos Administrativos, o nome da política era **Desabilitar envio quando as extensões da Web não podem ser carregadas**. Substitua esse nome em etapas posteriores, se necessário.

#### <a name="what-the-policy-does"></a>O que a política faz

Por motivos de conformidade, os administrador podem precisar garantir que os usuários não possam enviar itens de mensagem de reunião até que o último suplemento Ao enviar esteja disponível para execução. Os administradores devem habilitar o envio de blocos de política de grupo quando os **suplementos da Web** não puderem ser carregados para que todos os suplementos sejam atualizados do Exchange e estejam disponíveis para verificar se cada mensagem ou item de reunião atende às regras e regulamentos esperados no envio.

|Status da política|Resultado|
|---|---|
|Desabilitado|Os manifestos baixados no momento dos suplementos ao enviar (não necessariamente as versões mais recentes) são executados em itens de mensagem ou reunião que estão sendo enviados. Esse é o status/comportamento padrão.|
|Habilitado|Depois que os manifestos mais recentes dos suplementos ao enviar são baixados do Exchange, os suplementos são executados nos itens de mensagem ou reunião que estão sendo enviados. Caso contrário, o envio será bloqueado.|

#### <a name="manage-the-on-send-policy"></a>Gerenciar a política Ao enviar

Por padrão, a política Ao enviar está desabilitada. Os administradores podem habilitar a política ao enviar, garantindo que a configuração de política de grupo do usuário bloqueie o envio quando os **suplementos da Web** não puderem ser carregados estiver definido como **Habilitado**. Para desabilitar a política para um usuário, o administrador deve defini-la como **Desabilitada**. Para gerenciar essa configuração de política, você pode fazer o seguinte:

1. Baixe a [ferramenta de Modelos Administrativos](https://www.microsoft.com/download/details.aspx?id=49030) mais recente.
1. Abra o Editor Política de Grupo Local (**gpedit.msc**).
1. Navegue **até Modelos Administrativos de** > **Configuração de**  > **Usuário do Microsoft Outlook 2016** >  **Segurança** > **Central de Confiabilidade**.
1. Selecione o **envio de bloco quando os suplementos da Web não puderem carregar a configuração** .
1. Abra o link para configuração Editar política.
1. No envio **de bloco quando os suplementos da Web** não puderem carregar a janela de diálogo, selecione Habilitado  ou Desabilitado conforme apropriado e,  em seguida, selecione **OK** ou Aplicar para colocar a atualização em vigor.

### <a name="mac"></a>[Mac](#tab/unix)

Os suplementos para Outlook no Mac que usam o recurso Ao enviar devem ser executados para qualquer usuário que os tenha instalado. No entanto, se os usuários precisarem executar o suplemento para atender aos padrões de conformidade, a configuração de caixa de correio a seguir deverá ser aplicada ao computador de cada usuário. Esta configuração ou chave é compatível com CFPreference. Isso significa que é possível defini-la usando um software de gerenciamento empresarial para Mac, como o Jamf Pro.

||Valor|
|:---|:---|
|**Domínio**|com.microsoft.outlook|
|**Chave**|OnSendAddinsWaitForLoad|
|**DataType**|Booliano|
|**Valores possíveis**|falso (padrão)<br>verdadeiro|
|**Disponibilidade**|16.27|
|**Comentários**|Essa chave cria uma política de onSendMailbox.|

#### <a name="what-the-setting-does"></a>O que a configuração faz

Por motivos de conformidade, os administradores podem precisar garantir que os usuários não possam enviar itens de mensagem ou de reunião até que os suplementos estejam disponíveis para execução. Os administradores devem habilitar a chave **OnSendAddinsWaitForLoad** para que todos os suplementos sejam atualizados no Exchange e estejam disponíveis para verificar se cada item de mensagem ou de reunião atende às regras e normas esperadas ao enviar.

|Estado da chave|Resultado|
|---|---|
|falso|Os manifestos baixados no momento dos suplementos ao enviar (não necessariamente as versões mais recentes) são executados em itens de mensagem ou reunião que estão sendo enviados. Esse é o estado/comportamento padrão.|
|verdadeiro|Depois que os manifestos mais recentes dos suplementos ao enviar são baixados do Exchange, os suplementos são executados nos itens de mensagem ou reunião que estão sendo enviados. Caso contrário, o envio será bloqueado e **o botão** Enviar será desabilitado.|

---

## <a name="on-send-feature-scenarios"></a>Cenários do recurso Ao enviar

Veja a seguir os cenários com suporte e sem suporte para suplementos que usam o recurso Ao enviar.

### <a name="event-handlers-are-dynamically-defined"></a>Os manipuladores de eventos são definidos dinamicamente

Os manipuladores de eventos do suplemento devem ser definidos `Office.initialize` `Office.onReady()` pelo tempo ou chamados (para obter mais informações, consulte Inicialização de um suplemento do [Outlook](../develop/loading-the-dom-and-runtime-environment.md#startup-of-an-outlook-add-in) e inicializar seu suplemento do [Office](../develop/initialize-add-in.md)). Se o código do manipulador for definido dinamicamente por determinadas circunstâncias durante a inicialização, você deverá criar uma função de stub para chamar o manipulador quando ele estiver completamente definido. A função stub deve ser referenciada no **\<Event\>** atributo do `FunctionName` elemento do manifesto. Essa solução alternativa garante que o manipulador esteja definido e pronto para ser referenciado uma vez `Office.initialize` ou executado `Office.onReady()` .

Se o manipulador não for definido depois que o suplemento for inicializado, o remetente será notificado de que "A função de retorno de chamada está inacessível" por meio de uma barra de informações no item de email.

### <a name="user-mailbox-has-the-on-send-add-in-feature-enabled-but-no-add-ins-are-installed"></a>A caixa de correio do usuário tem o recurso de suplemento Ao enviar habilitado, mas nenhum suplemento está instalado

Nesse cenário, o usuário poderá enviar itens de reunião e mensagens sem nenhum suplemento em execução.

### <a name="user-mailbox-has-the-on-send-add-in-feature-enabled-and-add-ins-that-supports-on-send-are-installed-and-enabled"></a>A caixa de correio do usuário tem o recurso de suplemento Ao enviar habilitado, e os suplementos compatíveis com Ao enviar estão instalados e habilitados

Os suplementos serão executados durante o evento de envio, que em seguida permitirão ou impedirão o usuário de enviar.

### <a name="mailbox-delegation-where-mailbox-1-has-full-access-permissions-to-mailbox-2"></a>Delegação de caixa de correio, onde a caixa de correio 1 tem permissões de acesso total à caixa de correio 2

#### <a name="web-browser-classic-outlook"></a>Navegador da Web (Outlook clássico)

|Cenário|Recurso Ao enviar da caixa de correio 1|Recurso Ao enviar da caixa de correio 2|Sessão Web do Outlook (clássico)|Resultado|Com suporte?|
|:------------|:------------|:--------------------------|:---------|:-------------|:-------------|
|1|Habilitado|Habilitado|Nova sessão|A caixa de correio 1 não consegue enviar um item de mensagem ou de reunião da caixa de correio 2.|Não há suporte atualmente. Como alternativa, use o cenário 3.|
|2|Desabilitado|Habilitado|Nova sessão|A caixa de correio 1 não consegue enviar um item de mensagem ou de reunião da caixa de correio 2.|Não há suporte atualmente. Como alternativa, use o cenário 3.|
|3|Habilitado|Habilitado|Mesma sessão|Os suplementos Ao enviar atribuídos à caixa de correio 1 são executados ao enviar.|Com suporte.|
|4|Habilitado|Desabilitado|Nova sessão|Nenhum suplemento Ao envio é executado; item de mensagem ou de reunião é enviado.|Com suporte.|

#### <a name="web-browser-modern-outlook-windows-mac"></a>Navegador da Web (Outlook moderno), Windows, Mac

Para impor o Ao enviar, os administradores devem garantir que a política tenha sido habilitada nas duas caixas de correio. Para saber como dar suporte ao acesso delegado em um suplemento, consulte Habilitar pastas [compartilhadas e cenários de caixa de correio compartilhada](delegate-access.md).

### <a name="user-mailbox-with-on-send-add-in-featurepolicy-enabled-add-ins-that-support-on-send-are-installed-and-enabled-and-offline-mode-is-enabled"></a>Caixa de correio do usuário com recurso/política de suplemento Ao enviar habilitado, os suplementos com suporte à funcionalidade Ao enviar estão instalados e habilitados e o modo offline está habilitado

Os suplementos Ao enviar serão executados de acordo com o estado online do usuário, o back-end do suplemento e o Exchange.

#### <a name="users-state"></a>Estado do usuário

Os suplementos Ao enviar serão executados durante o envio se o usuário estiver online. Se o usuário estiver offline, os suplementos Ao enviar não serão executados e o item de mensagem ou de reunião não será enviado.

#### <a name="add-in-backends-state"></a>Estado do back-end do suplemento

Um suplemento Ao enviar será executado se o seu back-end estiver online e acessível. Se o back-end estiver offline, ao enviar será desabilitado.

#### <a name="exchanges-state"></a>Estado do Exchange

Os suplementos Ao enviar serão executados durante o envio se o servidor do Exchange estiver online e acessível. Se o suplemento Ao enviar não puder alcançar o Exchange e a política ou cmdlet aplicável estiverem ativados, o envio será desabilitado.

> [!NOTE]
> No Mac, em qualquer estado offline, o botão **Enviar** (ou o botão **Enviar Atualização** para reuniões existentes) está desabilitado e uma notificação é exibida informando que sua organização não permite envio quando o usuário está offline.

### <a name="user-can-edit-item-while-on-send-add-ins-are-working-on-it"></a>O usuário pode editar o item enquanto os suplementos ao enviar estão trabalhando nele

Enquanto suplementos ao enviar estão processando um item, o usuário pode editar o item adicionando, por exemplo, texto ou anexos inadequados. Se você quiser impedir que o usuário edite o item enquanto o suplemento estiver processando ao enviar, poderá implementar uma solução alternativa usando uma caixa de diálogo. Essa solução alternativa pode ser usada em Outlook na Web (clássico), Windows e Mac.

> [!IMPORTANT]
> Modo Outlook na Web: para impedir que o usuário edite o item enquanto o suplemento está processando no envio, você deve definir o sinalizador *OnSendAddinsEnabled* `true` como conforme descrito nos [suplementos instalar o Outlook](outlook-on-send-addins.md?tabs=modern#install-outlook-add-ins-that-use-on-send) que usam a seção ao enviar anteriormente neste artigo.

No manipulador ao enviar:

1. Chame [displayDialogAsync](/javascript/api/office/office.ui?view=outlook-js-preview&preserve-view=true#office-office-ui-displaydialogasync-member(1)) para abrir uma caixa de diálogo para que os cliques e pressionamentos de teclas do mouse sejam desabilitados.

    > [!IMPORTANT]
    > Para obter esse comportamento no Outlook na Web clássico, você deve definir a propriedade [displayInIframe](/javascript/api/office/office.dialogoptions?view=outlook-js-preview&preserve-view=true#office-office-dialogoptions-displayiniframe-member) como `true` no `options` parâmetro da `displayDialogAsync` chamada.

1. Implemente o processamento do item.
1. Feche a caixa de diálogo. Além disso, manipule o que acontece se o usuário fechar a caixa de diálogo.

## <a name="code-examples"></a>Exemplos de código

Os seguintes exemplos de código mostram como criar um suplemento simples Ao enviar. Para baixar o exemplo de código em que esses exemplos se baseiam, consulte [Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send).

> [!TIP]
> Se você usar uma caixa de diálogo com o evento ao enviar, feche a caixa de diálogo antes de concluir o evento.

### <a name="manifest-version-override-and-event"></a>Manifesto, versão de substituição e evento

Um exemplo de código [Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send) inclui dois manifestos:

- `Contoso Message Body Checker.xml` &ndash; Mostra como verificar se o corpo de uma mensagem apresenta palavras restritas ou informações confidenciais ao enviar.  

- `Contoso Subject and CC Checker.xml` &ndash; Mostra como adicionar um destinatário à linha CC e verifica se a mensagem inclui uma linha de assunto ao enviar.  

No arquivo de manifesto `Contoso Message Body Checker.xml`, inclua o arquivo de função e o nome da função que deve ser chamada no evento `ItemSend`. A operação é executada de maneira síncrona.

```xml
<Hosts>
    <Host xsi:type="MailHost">
        <DesktopFormFactor>
            <!-- The functionfile and function name to call on message send.  -->
            <!-- In this case, the function validateBody will be called within the JavaScript code referenced in residUILessFunctionFileUrl. -->
            <FunctionFile resid="residUILessFunctionFileUrl" />
            <ExtensionPoint xsi:type="Events">
                <Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="validateBody" />
            </ExtensionPoint>
        </DesktopFormFactor>
    </Host>
</Hosts>
```

> [!IMPORTANT]
> Se você estiver usando o Visual Studio 2019 para desenvolver seu suplemento ao enviar, poderá receber um aviso de validação como o seguinte: "Este é um xsi:type 'http://schemas.microsoft.com/office/mailappversionoverrides/1.1:Events'inválido". Para contornar isso, você precisará de uma versão mais recente do MailAppVersionOverridesV1_1.xsd, que foi fornecida como um gist do GitHub em um [blog sobre esse aviso](https://theofficecontext.com/2018/11/29/visual-studio-2017-this-is-an-invalid-xsitype-mailappversionoverrides-1-1event/).

Para o arquivo de manifesto `Contoso Subject and CC Checker.xml`, o exemplo a seguir mostra o arquivo de função e o nome da função para chamar o evento de envio de mensagem.

```xml
<Hosts>
    <Host xsi:type="MailHost">
        <DesktopFormFactor>
            <!-- The functionfile and function name to call on message send.  -->
            <!-- In this case the function validateSubjectAndCC will be called within the JavaScript code referenced in residUILessFunctionFileUrl. -->
            <FunctionFile resid="residUILessFunctionFileUrl" />
            <ExtensionPoint xsi:type="Events">
                <Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="validateSubjectAndCC" />
            </ExtensionPoint>
        </DesktopFormFactor>
    </Host>
</Hosts>
```

A API Ao enviar requer `VersionOverrides v1_1`. Veja a seguir como adicionar o nó `VersionOverrides` em seu manifesto.

```xml
 <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
     <!-- On-send requires VersionOverridesV1_1 -->
     <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
         ...
     </VersionOverrides>
</VersionOverrides>
```

> [!NOTE]
> Para obter mais informações, confira o seguinte:
>
> - [Manifestos de suplementos do Outlook](manifests.md)
> - [Manifesto XML dos Suplementos do Office](../develop/add-in-manifests.md)

### <a name="event-and-item-objects-and-bodygetasync-and-bodysetasync-methods"></a>Objetos `Event` e `item`, e os métodos `body.getAsync` e `body.setAsync`

Para acessar o item de mensagem ou de reunião selecionado no momento (neste exemplo, a mensagem redigida recentemente), use o namespace `Office.context.mailbox.item`. O evento `ItemSend` é passado automaticamente pelo recurso Ao enviar para a função especificada no manifesto&mdash;neste exemplo, a função `validateBody`.

```js
let mailboxItem;

Office.initialize = function (reason) {
    mailboxItem = Office.context.mailbox.item;
}

// Entry point for Contoso Message Body Checker add-in before send is allowed.
// <param name="event">ItemSend event is automatically passed by on-send code to the function specified in the manifest.</param>
function validateBody(event) {
    mailboxItem.body.getAsync("html", { asyncContext: event }, checkBodyOnlyOnSendCallBack);
}
```

A função `validateBody` obtém o corpo atual no formato especificado (HTML) e passa o objeto de evento `ItemSend` que o código deseja para acessar o método de retorno. Além do método `getAsync`, o objeto `Body` também fornece um método `setAsync` que você pode usar para substituir o corpo pelo texto especificado.

> [!NOTE]
> Para saber mais, confira [Objeto do Evento](/javascript/api/office/office.addincommands.event) e [Body.getAsync](/javascript/api/outlook/office.body#outlook-office-body-getasync-member(1)).
  
### <a name="notificationmessages-object-and-eventcompleted-method"></a>Objeto `NotificationMessages` e método `event.completed`

A função `checkBodyOnlyOnSendCallBack` usa uma expressão regular para determinar se o corpo da mensagem contém palavras bloqueadas. Se ela encontrar uma correspondência com uma matriz de palavras restritas, bloqueará os emails de serem enviados e notificará o remetente pela barra de informações. Para fazer isso, ele usa a propriedade `notificationMessages` do objeto `Item` para retornar um objeto `NotificationMessages`. Ele, em seguida, adiciona uma notificação ao item chamando o método `addAsync`, como mostrado no exemplo a seguir.

```js
// Determine whether the body contains a specific set of blocked words. If it contains the blocked words, block email from being sent. Otherwise allow sending.
// <param name="asyncResult">ItemSend event passed from the calling function.</param>
function checkBodyOnlyOnSendCallBack(asyncResult) {
    const listOfBlockedWords = new Array("blockedword", "blockedword1", "blockedword2");
    const wordExpression = listOfBlockedWords.join('|');

    // \b to perform a "whole words only" search using a regular expression in the form of \bword\b.
    // i to perform case-insensitive search.
    const regexCheck = new RegExp('\\b(' + wordExpression + ')\\b', 'i');
    const checkBody = regexCheck.test(asyncResult.value);

    if (checkBody) {
        mailboxItem.notificationMessages.addAsync('NoSend', { type: 'errorMessage', message: 'Blocked words have been found in the body of this email. Please remove them.' });
        // Block send.
        asyncResult.asyncContext.completed({ allowEvent: false });
    }

    // Allow send.
    asyncResult.asyncContext.completed({ allowEvent: true });
}
```

A seguir estão os parâmetros para o `addAsync` método.

- `NoSend` &ndash; uma cadeia de caractere que é uma chave especificada pelo desenvolvedor para fazer referência a uma mensagem de notificação. Você pode usá-la para modificar esta mensagem mais tarde. A chave não pode ter mais de 32 caracteres.
- `type` &ndash; uma das propriedades do parâmetro de objeto JSON. Representa o tipo de uma mensagem; os tipos correspondem aos valores da enumeração [Office.MailboxEnums.ItemNotificationMessageType](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype). Os valores possíveis são indicador de progresso, mensagem informativa ou mensagem de erro. Neste exemplo, `type` é uma mensagem de erro.  
- `message` &ndash; uma das propriedades do parâmetro de objeto JSON. Neste exemplo, `message` é o texto da mensagem de notificação.

Para sinalizar que o suplemento terminou de processar o evento `ItemSend` disparado pela operação enviar, chame o método `event.completed({allowEvent:Boolean})`. A propriedade `allowEvent` é um booleano. Se for definido como `true`, o envio será permitido. Se definido como `false`, a mensagem de email será impedida de ser enviada.

> [!NOTE]
> Para saber mais, confira [notificationMessages](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) e [completed](/javascript/api/office/office.addincommands.event).

### <a name="replaceasync-removeasync-and-getallasync-methods"></a>Métodos `replaceAsync`, `removeAsync` e `getAllAsync`

Além do método `addAsync`, o objeto `NotificationMessages` também inclui os métodos `replaceAsync`, `removeAsync` e `getAllAsync`.  Esses métodos não são usados neste exemplo de código.  Para saber mais, veja [NotificationMessages](/javascript/api/outlook/office.notificationmessages).

### <a name="subject-and-cc-checker-code"></a>Código do Assunto e do verificador de CC

O exemplo de código a seguir mostra como adicionar um destinatário à linha CC e verifica se a mensagem inclui um assunto ao enviar. Este exemplo usa o recurso Ao enviar para permitir ou proibir o envio de um email.  

```js
// Invoke by Contoso Subject and CC Checker add-in before send is allowed.
// <param name="event">ItemSend event is automatically passed by on-send code to the function specified in the manifest.</param>
function validateSubjectAndCC(event) {
    shouldChangeSubjectOnSend(event);
}

// Determine whether the subject should be changed. If it is already changed, allow send. Otherwise change it.
// <param name="event">ItemSend event passed from the calling function.</param>
function shouldChangeSubjectOnSend(event) {
    mailboxItem.subject.getAsync(
        { asyncContext: event },
        function (asyncResult) {
            addCCOnSend(asyncResult.asyncContext);
            //console.log(asyncResult.value);
            // Match string.
            const checkSubject = (new RegExp(/\[Checked\]/)).test(asyncResult.value)
            // Add [Checked]: to subject line.
            subject = '[Checked]: ' + asyncResult.value;

            // Determine whether a string is blank, null, or undefined.
            // If yes, block send and display information bar to notify sender to add a subject.
            if (asyncResult.value === null || (/^\s*$/).test(asyncResult.value)) {
                mailboxItem.notificationMessages.addAsync('NoSend', { type: 'errorMessage', message: 'Please enter a subject for this email.' });
                asyncResult.asyncContext.completed({ allowEvent: false });
            }
            else {
                // If can't find a [Checked]: string match in subject, call subjectOnSendChange function.
                if (!checkSubject) {
                    subjectOnSendChange(subject, asyncResult.asyncContext);
                    //console.log(checkSubject);
                }
                else {
                    // Allow send.
                    asyncResult.asyncContext.completed({ allowEvent: true });
                }
            }
        });
}

// Add a CC to the email. In this example, CC contoso@contoso.onmicrosoft.com
// <param name="event">ItemSend event passed from calling function</param>
function addCCOnSend(event) {
    mailboxItem.cc.setAsync(['Contoso@contoso.onmicrosoft.com'], { asyncContext: event });
}

// Determine whether the subject should be changed. If it is already changed, allow send, otherwise change it.
// <param name="subject">Subject to set.</param>
// <param name="event">ItemSend event passed from the calling function.</param>
function subjectOnSendChange(subject, event) {
    mailboxItem.subject.setAsync(
        subject,
        { asyncContext: event },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                mailboxItem.notificationMessages.addAsync('NoSend', { type: 'errorMessage', message: 'Unable to set the subject.' });

                // Block send.
                asyncResult.asyncContext.completed({ allowEvent: false });
            }
            else {
                // Allow send.
                asyncResult.asyncContext.completed({ allowEvent: true });
            }
        });
}
```

Para saber mais sobre como adicionar um destinatário à linha CC e verificar se a mensagem de e-mail inclui uma linha de assunto ao enviar e para ver as APIs que você pode usar, consulte o [exemplo Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send). O código é bem comentado.

## <a name="debug-outlook-add-ins-that-use-on-send"></a>Depurar suplementos do Outlook que usam ao enviar

Para obter instruções sobre como depurar seu suplemento ao enviar, consulte comandos [de função de depuração em suplementos do Outlook](debug-ui-less.md).

> [!TIP]
> Se o erro "A função de retorno de chamada estiver inacessível" aparece quando os usuários executam o suplemento e o manipulador de eventos do suplemento é definido dinamicamente, você deve criar uma função de stub como uma solução alternativa. Consulte [Manipuladores de eventos são definidos dinamicamente](#event-handlers-are-dynamically-defined) para obter mais informações.

## <a name="see-also"></a>Confira também

- [Visão geral da arquitetura e dos recursos de suplementos do Outlook](outlook-add-ins-overview.md)
- [Suplemento do Outlook para demonstração de comando de suplemento](https://github.com/OfficeDev/outlook-add-in-command-demo)
