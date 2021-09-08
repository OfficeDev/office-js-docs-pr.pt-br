---
title: Solução de problemas de ativação de suplementos contextuais do Outlook
description: Possíveis motivos para o seu complemento não ser ativado como você espera.
ms.date: 09/02/2020
localization_priority: Normal
ms.openlocfilehash: 7999bac00e2d3b1b7ecb3f58e6d70a247500982e
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58938068"
---
# <a name="troubleshoot-outlook-add-in-activation"></a>Solução de problemas de ativação de suplementos do Outlook

Outlook ativação de um complemento contextual se baseia nas regras de ativação no manifesto do complemento. Quando as condições para o item selecionado no momento atendem às regras de ativação do complemento, o aplicativo ativa e exibe o botão de complemento na interface do usuário do Outlook (painel de seleção de complementos para complementos de composição, barra de complementos para complementos de leitura). No entanto, se seu suplemento não for ativado conforme o esperado, procure a causa nas áreas a seguir.

## <a name="is-user-mailbox-on-a-version-of-exchange-server-that-is-at-least-exchange-2013"></a>A caixa de correio do usuário está em uma versão do Exchange Server que tenha pelo menos o Exchange 2013?

Primeiro, verifique se a conta de email do usuário que sendo testada está uma versão do Exchange Server que tenha pelo menos o Exchange 2013. Se você estiver usando recursos específicos lançados após o Exchange 2013, verifique se que a conta do usuário está na versão adequada do Exchange.

Você pode verificar a versão do Exchange 2013 usando uma das seguintes abordagens.

- Verifique com o administrador do Exchange Server.

- Se você estiver testando o suplemento no Outlook na Web ou em dispositivos móveis em um depurador de script (por exemplo, o Depurador JScript que acompanha o Internet Explorer), procure o atributo **src** da marca **script** que especifica o local do qual os scripts são carregados. O caminho deve conter uma subcadeia de caracteres **owa/15.0.516.x/owa2/...**, em que **15.0.516.x** representa a versão do Exchange Server, como **15.0.516.2**.

- Como alternativa, você pode usar a propriedade [Office.context.mailbox.diagnostics.hostVersion](/javascript/api/outlook/office.diagnostics#hostVersion) para verificar a versão. No Outlook na Web e nos dispositivos móveis, essa propriedade retorna a versão do Exchange Server.

- Se você puder testar o Outlook do Outlook, poderá usar a seguinte técnica de depuração simples que usa o modelo de objeto Outlook e o editor Visual Basic.

    1. Primeiro, verifique se as macros estão habilitadas para o Outlook. Escolha **Arquivo**, **Opções**, **Central de Confiabilidade**, **Configurações da Central de Confiabilidade**, **Configurações de Macro**. Verifique se a opção **Notificações para todas as macros** está selecionada na Central de Confiabilidade. Você deve escolher também **Habilitar Macros**, durante a inicialização do Outlook.

    1. Na guia **Desenvolvedor** da faixa de opções, escolha **Visual Basic**.

       > [!NOTE]
       > Não consegue ver a guia **Desenvolvedor**? Confira [Como Mostrar a Guia Desenvolvedor na Faixa de Opções](/visualstudio/vsto/how-to-show-the-developer-tab-on-the-ribbon) para ativá-la.

    1. No Editor do Visual Basic, escolha **Exibir**, **Janela Imediata**.

    1. Digite o texto a seguir na janela Imediata para exibir a versão do Exchange Server. A versão principal do valor retornado deve ser igual ou maior que 15.

       - Se houver apenas uma conta do Exchange no perfil do usuário:

       ```vb
        ?Session.ExchangeMailboxServerVersion
       ```

       - Caso haja várias contas do Exchange no mesmo perfil de usuário (`emailAddress` representa uma cadeia de caracteres que contém o endereço SMTP principal do usuário):

       ```vb
        ?Session.Accounts.Item(emailAddress).ExchangeMailboxServerVersion
       ```

## <a name="is-the-add-in-disabled"></a>O suplemento está desabilitado?

Qualquer um dos clientes avançados do Outlook pode desativar um suplemento por motivos de desempenho, incluindo uso excedido dos limites de memória ou núcleo da CPU, tolerância a falhas e período de tempo para processar todas as expressões regulares de um suplemento. Quando isso acontece, o cliente avançado do Outlook exibe uma notificação de que vai desabilitar o suplemento.

> [!NOTE]
> Somente os clientes avançados do Outlook monitoram o uso do recurso, mas desabilitar um suplemento em um cliente avançado do Outlook também desabilita o suplemento no Outlook na Web e nos dispositivos móveis.

Use uma das seguintes abordagens para verificar se um complemento está desabilitado.

- No Outlook na Web, entre diretamente na conta de email, escolha o ícone Configurações e **Gerenciar Suplementos** para ir para o Centro de Administração do Exchange, onde você pode verificar se o suplemento está habilitado.

- No Outlook no Windows, vá para o modo de exibição Backstage e escolha **Gerenciar suplementos**. Entre no Centro de Administração do Exchange para verificar se o suplemento foi habilitado.

- No Outlook no Mac, escolha **Gerenciar suplementos** na barra do suplemento. Entre no Centro de Administração do Exchange para verificar se o suplemento foi habilitado.

## <a name="does-the-tested-item-support-outlook-add-ins-is-the-selected-item-delivered-by-a-version-of-exchange-server-that-is-at-least-exchange-2013"></a>O item testado dá suporte a suplementos do Outlook? O item selecionado foi fornecido por uma versão do Exchange Server que tenha pelo menos o Exchange 2013?

Se o seu suplemento do Outlook é um suplemento de leitura e deve ser ativado quando o usuário está exibindo uma mensagem (inclusive mensagens de email, solicitações de reunião, respostas e cancelamentos de reunião) ou um compromisso, embora esses itens geralmente sejam compatíveis com suplementos, há exceções. Verifique se o item selecionado é um dos [listados em que os suplementos do Outlook não são ativados](outlook-add-ins-overview.md#mailbox-items-available-to-add-ins).

Além disso, já que os compromissos sempre são salvos no formato Rich Text, uma regra [ItemHasRegularExpressionMatch](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule) que especifica um valor **PropertyName** de **BodyAsHTML** não ativaria um suplemento em um compromisso ou uma mensagem salva em texto sem formatação ou no formato Rich Text.

Mesmo se um item de email não for um dos tipos de acima, se ele não veio de uma versão do Exchange Server com no mínimo o Exchange 2013, as propriedades e entidades conhecidas, como o endereço SMTP do remetente, não seriam identificadas no item. Regras de ativação que dependem dessas entidades ou propriedades não seriam atendidas e o suplemento não seria ativado.

Se seu suplemento é um suplemento de redação e deve ser ativado quando o usuário cria uma mensagem ou solicitação de reunião, verifique se o item não está protegido por IRM. No entanto, há algumas exceções.

1. Os suplementos são ativados em mensagens assinadas digitalmente no Outlook associadas a uma assinatura do Microsoft 365. No Windows, esse suporte foi introduzido com a compilação 8711.1000.
1. A partir do Outlook, build 13229.10000, no Windows, os suplementos agora podem ser ativados nos itens protegidos por IRM.  Para obter mais informações sobre esse suporte na visualização, consulte Ativação de complemento em itens protegidos [pelo IRM (Gerenciamento](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md#add-in-activation-on-items-protected-by-information-rights-management-irm)de Direitos de Informação).

## <a name="is-the-add-in-manifest-installed-properly-and-does-outlook-have-a-cached-copy"></a>O manifesto do suplemento está instalado corretamente? O Outlook tem uma cópia armazenada em cache?

Esse cenário se aplica apenas Outlook em Windows. Normalmente, quando você instala um suplemento do Outlook para uma caixa de correio, o Exchange Server copia manifesto do suplemento do local indicado para a caixa de correio nesse Exchange Server. Sempre que Outlook, ele lê todos os manifestos instalados para essa caixa de correio em um cache temporário no local a seguir.

```text
%LocalAppData%\Microsoft\Office\16.0\WEF
```

Por exemplo, para o usuário John, o cache pode estar em C:\Users\john\AppData\Local\Microsoft\Office\16.0\WEF.

> [!IMPORTANT]
> Para Outlook 2013 no Windows, use 15.0 em vez de 16,0 para que o local seja:
>
> ```text
> %LocalAppData%\Microsoft\Office\15.0\WEF
> ```

Se um suplemento não foi ativado para todos os itens, o manifesto talvez não tenha sido instalado corretamente no Exchange Server ou o Outlook não leu o manifesto corretamente na inicialização. Usando o Centro de Administração do Exchange, verifique se o suplemento está instalado e habilitado para sua caixa de correio e reinicie o Exchange Server, se necessário.

A Figura 1 mostra um resumo das etapas para verificar se o Outlook tem uma versão válida do manifesto.

**Figura 1. Fluxograma das etapas para verificar se o Outlook armazenou o manifesto em cache adequadamente**

![Flow gráfico para verificar manifesto.](../images/troubleshoot-manifest-flow.png)

O procedimento a seguir descreve os detalhes.

1. Se você modificou o manifesto enquanto o Outlook estava aberto e não estava usando o Visual Studio 2012 ou uma versão posterior do Visual Studio para desenvolver o suplemento, deve desinstalar o suplemento e reinstalá-lo usando o Centro de Administração do Exchange.

1. Reinicie o Outlook e teste se ele agora ativa o suplemento.

1. Se o Outlook não ativar o suplemento, verifique se tem uma cópia corretamente armazenada em cache do manifesto para o suplemento. Procure no caminho a seguir.

    ```text
    %LocalAppData%\Microsoft\Office\16.0\WEF
    ```

    Você pode encontrar o manifesto na subpasta a seguir.

    ```text
    \<insert your guid>\<insert base 64 hash>\Manifests\<ManifestID>_<ManifestVersion>
    ```

    > [!NOTE]
    > A seguir está um exemplo de um caminho para um manifesto instalado para uma caixa de correio para o usuário John.
    >
    > ```text
    > C:\Users\john\appdata\Local\Microsoft\Office\16.0\WEF\{8D8445A4-80E4-4D6B-B7AC-D4E6AF594E73}\GoRshCWa7vW8+jhKmyiDhA==\Manifests\b3d7d9d5-6f57-437d-9830-94e2aaccef16_1.2
    > ```

    Verifique se o manifesto do suplemento que você está testando está entre os manifestos armazenados em cache.

1. Se o manifesto estiver no cache, ignore o restante desta seção e considere outros motivos possíveis após esta seção.

1. Se o manifesto não estiver no cache, verifique se o Outlook leu o manifesto do Exchange Server com êxito. Para fazer isso, use o Visualizador de Eventos do Windows:

    1. Em **Logs do Windows**, escolha **Aplicativo**.

    1. Procure um evento razoavelmente recente com ID de Evento igual a 63, que representa o Outlook baixando um manifesto de um Exchange Server.

    1. Se Outlook ler um manifesto com êxito, o evento registrado deverá ter a seguinte descrição.

        ```text
        The Exchange web service request GetAppManifests succeeded.
        ```

        Ignore o restante desta seção e considere outros motivos possíveis após esta seção.

1. Se você não vir um evento bem-sucedido, feche Outlook e exclua todos os manifestos no caminho a seguir.

    ```text
    %LocalAppData%\Microsoft\Office\16.0\WEF\<insert your guid>\<insert base 64 hash>\Manifests\
    ```

    Inicie o Outlook e teste se ele agora ativa o suplemento.

1. Se o Outlook não ativar o suplemento, volte para a etapa 3 e verifique novamente se o Outlook leu o manifesto corretamente.

## <a name="is-the-add-in-manifest-valid"></a>O manifesto do suplemento é válido?

Confira [Validar e solucionar problemas com seu manifesto](../testing/troubleshoot-manifest.md) para depurar problemas do manifesto de suplemento.

## <a name="are-you-using-the-appropriate-activation-rules"></a>Você está usando as regras de ativação apropriadas?

A partir da versão 1.1 do esquema de manifestos dos suplementos do Office, é possível criar suplementos que são ativados quando o usuário está em um formulário de redação (suplementos de redação) ou em um formulário de leitura (suplementos de leitura). Não deixe de especificar as regras de ativação apropriadas para cada tipo de formulário em que seu suplemento deve ser ativado. Por exemplo, você pode ativar suplementos de redação usando apenas regras [ItemIs](../reference/manifest/rule.md#itemis-rule) com o atributo **FormType** definido como **Edit** ou **ReadOrEdit** e não usar qualquer dos outros tipos de regras, como [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) e [ItemHasRegularExpressionMatch](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule), para suplementos de redação. Para saber mais, confira [Regras de ativação para suplementos do Outlook](activation-rules.md).

## <a name="if-you-use-a-regular-expression-is-it-properly-specified"></a>Se você está usando uma expressão regular, ela foi especificada corretamente?

Já que as expressões regulares nas regras de ativação fazem parte do arquivo de manifesto XML para um suplemento de leitura, se uma expressão regular usa determinados caracteres, use a sequência de escape correspondente compatível com processadores XML. A Tabela 1 lista esses caracteres especiais.

**Tabela 1. Sequências de escape para expressões regulares**

|**Caractere**|**Descrição**|**Sequência de escape a ser usada**|
|:-----|:-----|:-----|
|`"`|Aspas duplas|&amp;quot;|
|`&`|E comercial|&amp;amp;|
|`'`|Apóstrofo|&amp;apos;|
|`<`|Sinal menor que|&amp;lt;|
|`>`|Sinal maior que|&amp;gt;|

## <a name="if-you-use-a-regular-expression-is-the-read-add-in-activating-in-outlook-on-the-web-or-mobile-devices-but-not-in-any-of-the-outlook-rich-clients"></a>Se você usa uma expressão regular, o suplemento de leitura está sendo ativado no Outlook na Web ou nos dispositivos móveis, mas não em clientes avançados do Outlook?

Os clientes avançados do Outlook usam um mecanismo de expressões regulares diferente daquele usado pelo Outlook na Web e pelos dispositivos móveis. Clientes avançados do Outlook usam o mecanismo de expressões regulares C++ fornecido como parte da biblioteca de modelo padrão do Visual Studio. Esse mecanismo é compatível com as normas ECMAScript 5. O Outlook na Web e os dispositivos móveis usam a avaliação da expressão regular que faz parte do JavaScript, é fornecida pelo navegador e dá suporte a um subconjunto dos ECMAScript 5.

Embora, na maioria dos casos, Outlook clientes encontrem as mesmas combinações para a mesma expressão regular em uma regra de ativação, há exceções. Por exemplo, se o regex incluir uma classe de caractere personalizada com base em classes de caracteres predefinidas, um cliente Outlook cliente rico poderá retornar resultados diferentes de Outlook na Web e dispositivos móveis. Por exemplo, classes de caracteres que contêm classes de caracteres abreviadas `[\d\w]` dentro delas retornam resultados diferentes. Nesse caso, para evitar resultados diferentes em diferentes aplicativos, use `(\d|\w)` em vez disso.

Teste sua expressão regular minuciosamente. Se ela retornar resultados diferentes, reescreva a expressão regular para ficar compatível em ambos os mecanismos. Para verificar os resultados de avaliação em um cliente avançado do Outlook, escreva um programa C++ pequeno que aplica a expressão regular em uma amostra do texto que você está tentando corresponder. Sendo executado no Visual Studio, o programa de teste C++ usaria a biblioteca de modelo padrão, simulando o comportamento do cliente avançado do Outlook ao executar a mesma expressão regular. Para verificar os resultados de avaliação do Outlook na Web ou nos dispositivos móveis, use seu avaliador de expressão regular JavaScript favorito.

## <a name="if-you-use-an-itemis-itemhasattachment-or-itemhasregularexpressionmatch-rule-have-you-verified-the-related-item-property"></a>Se você usa uma regra ItemIs, ItemHasAttachment ou ItemHasRegularExpressionMatch, já verificou a propriedade do item relacionado?

Se você usa uma regra de ativação **ItemHasRegularExpressionMatch**, verifique se o valor do atributo **PropertyName** é o que você espera do item selecionado. Veja a seguir algumas dicas para depurar as propriedades correspondentes.

- Se o item selecionado for uma mensagem e especificar **BodyAsHTML** no atributo **PropertyName**, abra a mensagem e escolha **Exibir Código-fonte** para verificar o corpo da mensagem na representação HTML desse item.

- Se o item selecionado for um compromisso ou se a regra de ativação especificar **BodyAsPlaintext** no **PropertyName**, você poderá usar o modelo de objeto do Outlook e o Editor do Visual Basic no Outlook no Windows:

    1. Verifique se as macros estão habilitadas e se a guia **Desenvolvedor** é exibida na faixa de opções do Outlook.

    1. No Editor do Visual Basic, escolha **Exibir**, **Janela Imediata**.

    1. Digite o texto a seguir para exibir várias propriedades dependendo do cenário.

        - O corpo HTML do item selecionado de compromisso ou mensagem no Outlook Explorer:

        ```vb
        ?ActiveExplorer.Selection.Item(1).HTMLBody
        ```
        - O corpo de texto sem formatação do item selecionado de compromisso ou mensagem no Outlook Explorer:

        ```vb
        ?ActiveExplorer.Selection.Item(1).Body
        ```
        - O corpo HTML do item selecionado de compromisso ou mensagem aberto no inspetor atual do Outlook:

        ```vb
        ?ActiveInspector.CurrentItem.HTMLBody
        ```
        - O corpo de texto sem formatação do item selecionado de compromisso ou mensagem aberto no inspetor atual do Outlook:

        ```vb
        ?ActiveInspector.CurrentItem.Body
        ```

Se a regra de ativação **ItemHasRegularExpressionMatch** especificar **Subject** ou **SenderSMTPAddress**, ou se você usar uma regra **ItemIs** ou **ItemHasAttachment** e conhecer ou quiser usar MAPI, pode usar [MFCMAPI](https://github.com/stephenegriffin/mfcmapi) para verificar o valor na Tabela 2 do qual a sua regra depende.

**Tabela 2. Regras de ativação e propriedades MAPI correspondentes**

|Tipo de regra|Verifique essa propriedade MAPI|
|:-----|:-----|
|Regra **ItemHasRegularExpressionMatch** com **Subject**|[PidTagSubject](/office/client-developer/outlook/mapi/pidtagsubject-canonical-property)|
|Regra **ItemHasRegularExpressionMatch** com **SenderSMTPAddress**|[PidTagSenderSmtpAddress](/office/client-developer/outlook/mapi/pidtagsendersmtpaddress-canonical-property) e [PidTagSentRepresentingSmtpAddress](/office/client-developer/outlook/mapi/pidtagsentrepresentingsmtpaddress-canonical-property)|
|**ItemIs**|[PidTagMessageClass](/office/client-developer/outlook/mapi/pidtagmessageclass-canonical-property)|
|**ItemHasAttachment**|[PidTagHasAttachments](/office/client-developer/outlook/mapi/pidtaghasattachments-canonical-property)|

Depois de verificar o valor da propriedade, você pode usar uma ferramenta de avaliação da expressão regular para testar se a expressão regular localiza uma correspondência a esse valor.

## <a name="does-outlook-apply-all-the-regular-expressions-to-the-portion-of-the-item-body-as-you-expect"></a>O Outlook aplica todas as expressões regulares à parte do corpo do item como você espera?

Esta seção aplica-se a todas as regras de ativação que usam expressões regulares, particularmente àquelas que serão aplicadas ao corpo do item, que pode ser grande e levar mais tempo para avaliar correspondências. Você deve estar ciente de que, mesmo que a propriedade de item da qual uma regra de ativação dependa tenha o valor esperado, Outlook pode não ser capaz de avaliar todas as expressões regulares em todo o valor da propriedade item. Para fornecer um desempenho razoável e controlar o uso excessivo de recursos por um complemento de leitura, Outlook observa os seguintes limites sobre o processamento de expressões regulares em regras de ativação em tempo de execução.

- O tamanho do corpo do item avaliado -- Há limites para a parte de um corpo de item no qual Outlook avalia uma expressão regular. Esses limites dependem do cliente Outlook, fator de formulário e formato do corpo do item. Confira os detalhes na Tabela 2 em [Limites de ativação e API JavaScript para suplementos do Outlook](limits-for-activation-and-javascript-api-for-outlook-add-ins.md).

- Número de correspondências de expressão regular: os clientes avançados do Outlook, o Outlook na Web e nos dispositivos móveis retornam, cada um, no máximo 50 correspondências de expressões regulares. Essas correspondências são exclusivas e correspondências duplicadas não contam para esse limite. Não suponha uma ordem nas correspondências retornadas e não suponha que a ordem em um cliente avançado do Outlook é a mesma no Outlook na Web e no OWA para Dispositivos. Se espera muitas correspondências para expressões regulares em suas regras de ativação e está faltando uma correspondência, é possível que você esteja excedendo esse limite.

- Comprimento de uma combinação de expressão regular - Há limites para o comprimento de uma expressão regular corresponder a que o Outlook aplicativo retornaria. Outlook não inclui nenhuma combinação acima do limite e não exibe nenhuma mensagem de aviso. Você pode executar sua expressão regular usando outras ferramentas de avaliação de regex ou um programa de teste C++ autônomo para verificar se há uma correspondência que excede esses limites. A Tabela 3 resume os limites. Para saber mais, confira a Tabela 3 em [Limites de ativação e API JavaScript para suplementos do Outlook](limits-for-activation-and-javascript-api-for-outlook-add-ins.md).

    **Tabela 3. Limites de comprimento para correspondência de uma expressão regular**

    |Limite de comprimento de uma correspondência de regex|Clientes avançados do Outlook|Outlook na Web ou em dispositivos móveis|
    |:-----|:-----|:-----|
    |O corpo do item é texto sem formatação|1,5 KB|3 KB|
    |Corpo do item é HTML|3 KB|3 KB|

- Tempo gasto na avaliação de todas as expressões regulares de um suplemento de leitura para um cliente avançado do Outlook: Por padrão, para cada suplemento de leitura, o Outlook deve concluir a avaliação de todas as expressões regulares em suas regras de ativação em um segundo. Caso contrário, o Outlook tenta mais três vezes e desabilita o suplemento se não conseguir concluir a avaliação. O Outlook exibe uma mensagem na barra de notificações de que o suplemento foi desabilitado. A quantidade de tempo disponível para sua expressão regular pode ser modificada com a definição de uma política de grupo ou uma chave do registro. 

   > [!NOTE]
   > Se o cliente avançado do Outlook desabilita um suplemento de leitura, o suplemento de leitura não fica disponível para uso na mesma caixa de correio no cliente avançado do Outlook, no Outlook na Web e nos dispositivos móveis.

## <a name="see-also"></a>Confira também

- [Implantar e instalar suplementos do Outlook para teste](testing-and-tips.md)
- [Regras de ativação para suplementos do Outlook](activation-rules.md)
- [Usar regras de ativação de expressões regulares para mostrar um suplemento do Outlook](use-regular-expressions-to-show-an-outlook-add-in.md)
- [Limites de ativação e API JavaScript para suplementos do Outlook](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
- [Validar e solucionar problemas com seu manifesto](../testing/troubleshoot-manifest.md)
