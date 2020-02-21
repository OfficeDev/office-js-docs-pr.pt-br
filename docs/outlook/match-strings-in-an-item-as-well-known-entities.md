---
title: Corresponder cadeias de caracteres como entidades conhecidas em um suplemento do Outlook
description: Usando a API JavaScript para Office, você pode obter cadeias de caracteres que correspondem a entidades conhecidas específicas para processá-las posteriormente.
ms.date: 04/15/2019
localization_priority: Normal
ms.openlocfilehash: 9ea34c53bd7c4c28ab5910b618c828ec59c3be92
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165804"
---
# <a name="match-strings-in-an-outlook-item-as-well-known-entities"></a>Corresponder cadeias de caracteres em um item do Outlook como entidades conhecidas

Antes de enviar um item de mensagem ou de solicitação de reunião, o Exchange Server analisa o conteúdo do item, identifica e apresenta determinadas cadeias de caracteres no assunto e no corpo semelhantes a entidades conhecidas do Exchange, como endereços de email, números de telefone e URLs. As mensagens e solicitações de reunião são fornecidas pelo Exchange Server em uma Caixa de Entrada do Outlook com entidades conhecidas carimbadas. 

Usando a API JavaScript para Office, você pode obter essas cadeias de caracteres que correspondem a entidades conhecidas específicas para processá-las posteriormente. Também pode especificar uma entidade conhecida em uma regra no manifesto do suplemento, para que o Outlook possa ativar o suplemento quando o usuário estiver exibindo um item que contém correspondências para essa entidade. Em seguida, é possível extrair e agir em relação às correspondências da entidade. 

Convém ser capaz de identificar ou extrair tais instâncias de uma mensagem ou compromisso selecionado. Por exemplo, você pode compilar um serviço de pesquisa invertida de telefones como um suplemento do Outlook. O suplemento pode extrair cadeias de caracteres no corpo ou assunto do item que se parecem com um número de telefone, fazer uma pesquisa invertida e exibir o proprietário registrado de cada número de telefone.

Este tópico apresenta essas entidades conhecidas, mostra exemplos de regras de ativação baseadas em entidades conhecidas e como extrair correspondências de entidade independentemente de ter usado entidades em regras de ativação.


## <a name="support-for-well-known-entities"></a>Suporte para entidades conhecidas

O Exchange Server carimba entidades conhecidas em um item de mensagem ou de solicitação de reunião depois que o remetente envia o item e antes de o Exchange entregar o item ao destinatário. Portanto, somente os itens que passaram pelo transporte do Exchange são carimbados, e o Outlook pode ativar suplementos com base nesses carimbos quando o usuário está exibindo esses itens. Do contrário, quando o usuário está redigindo ou visualizando um item que está na pasta Itens Enviados, como o item não passou por transporte, o Outlook não pode ativar suplementos com base em entidades conhecidas. 

Da mesma forma, você não pode extrair entidades conhecidas em itens que estão sendo redigidos ou estão na pasta Itens Enviados, já que esses itens não passaram pelo transporte e não foram carimbados. Para saber mais sobre os tipos de itens que dão suporte à ativação, confira [Regras de ativação para suplementos do Outlook](activation-rules.md).

A tabela a seguir lista as entidades que têm suporte e são reconhecidas pelo Exchange Server e pelo Outlook (por isso chamadas "entidades conhecidas") e o tipo de objeto de uma instância de cada entidade. O reconhecimento de linguagem natural de uma cadeia de caracteres como uma dessas entidades baseia-se em um modelo de aprendizagem que foi treinado com grande quantidade de dados. Portanto, o reconhecimento é não determinístico. Confira [Dicas para usar entidades conhecidas](#tips-for-using-well-known-entities) a fim de saber mais sobre condições de reconhecimento.

**Tabela 1. Entidades compatíveis e os respectivos tipos**

|Tipo de entidade|Condições de reconhecimento|Tipo de objeto|
|:-----|:-----|:-----|
|**Endereço**|Endereços nos Estados Unidos. Por exemplo: 1234 Main Street, Redmond, WA 07722. Normalmente, para um endereço ser reconhecido, ele deve seguir a estrutura de um endereço postal dos Estados Unidos, com a maioria dos elementos de nome da rua, número, cidade, estado e CEP. O endereço pode ser especificado em uma ou várias linhas.|Objeto JavaScript **String**|
|**Contato**|Uma referência às informações de uma pessoa assim reconhecida em sua língua materna. O reconhecimento de um contato depende do contexto. Por exemplo, uma assinatura no final de uma mensagem ou o nome da pessoa que aparece perto de algumas das seguintes informações: número de telefone, endereço, endereço de e-mail e URL.|Objeto [Contact](/javascript/api/outlook/office.contact)|
|**EmailAddress**|Endereços de email SMTP.|Objeto JavaScript **String**|
|**MeetingSuggestion**|Uma referência a uma reunião ou a um evento. Por exemplo, o Exchange 2013 reconheceria o seguinte texto como uma sugestão de reunião:  _Vamos marcar um almoço amanhã._|Objeto [MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|
|**PhoneNumber**|Números de telefone dos Estados Unidos. Por exemplo:  _(235) 555-0110_|Objeto [PhoneNumber](/javascript/api/outlook/office.phonenumber)|
|**TaskSuggestion**|Frases acionáveis em um email. Por exemplo:  _Atualize a planilha._|Objeto [TaskSuggestion](/javascript/api/outlook/office.tasksuggestion)|
|**Url**|Um endereço Web que especifica explicitamente o local de rede e o identificador de um recurso da Web. O Exchange Server não requer o protocolo de acesso no endereço Web e não reconhece URLs que são inseridas no texto do link como instâncias da entidade **Url**. O Exchange Server pode corresponder aos seguintes exemplos `www.youtube.com/user/officevideos` :`https://www.youtube.com/user/officevideos` |Objeto JavaScript **String**|

<br/>

A figura a seguir descreve como o Exchange Server e o Outlook dão suporte a entidades conhecidas para suplementos e o que estes podem fazer com entidades conhecidas. Confira [Recuperar entidades em seu suplemento](#retrieving-entities-in-your-add-in) e [Ativar um suplemento com base na existência de uma entidade](#activating-an-add-in-based-on-the-existence-of-an-entity) para obter mais detalhes sobre como usar essas entidades.

**Como o Exchange Server, o Outlook e os suplementos dão suporte a entidades conhecidas**

![Support and use of well-known entities in mail app](../images/well-known-entities-info.png)


## <a name="permissions-to-extract-entities"></a>Permissões para extrair entidades

Para extrair entidades no seu código JavaScript ou fazer com que seu suplemento seja ativado com base na existência de determinadas entidades conhecidas, verifique se você solicitou as permissões apropriadas no manifesto do suplemento.

A especificação da permissão restrita padrão permite que o suplemento extraia as entidades **Address**, **MeetingSuggestion** ou **TaskSuggestion**. Para extrair as outras entidades, especifique as permissões de leitura de item, leitura/gravação de item ou leitura/gravação de caixa de correio. Para fazer isso no manifesto, use o elemento [Permissions](../reference/manifest/permissions.md) e especifique a permissão apropriada &mdash; **Restricted**, **ReadItem**, **ReadWriteItem** ou **ReadWriteMailbox** &mdash; como no exemplo abaixo:

```xml
<Permissions>ReadItem</Permissions>
```


## <a name="retrieving-entities-in-your-add-in"></a>Recuperar entidades no seu suplemento

Desde que o assunto ou corpo do item que está sendo visualizado pelo usuário contenha cadeias de caracteres que o Exchange e o Outlook possam reconhecer como entidades conhecidas, essas instâncias estarão disponíveis para suplementos. Elas estarão disponíveis mesmo que um suplemento não seja ativado com base em entidades conhecidas. Com a permissão apropriada, você pode usar os métodos **getEntities** ou **getEntitiesByType** para recuperar entidades conhecidas que estejam presentes na mensagem ou compromisso atual.

O método **getEntities** retorna uma matriz de objetos [Entities](/javascript/api/outlook/office.entities) que contém todas as entidades conhecidas no item.

Se você estiver interessado em um determinado tipo de entidades, use o método **getEntitiesByType**, que retorna uma matriz somente com as entidades desejadas. A enumeração [EntityType](/javascript/api/outlook/office.mailboxenums.entitytype) representa todos os tipos de entidades conhecidas que você pode extrair.

Após chamar **getEntities**, você pode usar a propriedade correspondente do objeto **Entities** para obter uma matriz de instâncias de um tipo de entidade. Dependendo do tipo de entidade, as instâncias na matriz podem ser apenas cadeias de caracteres ou podem mapear para objetos específicos. 

Como o exemplo mostrado na figura anterior, acesse a matriz retornada por `getEntities().addresses[]` para obter endereços no item. A propriedade **Entities.addresses** retorna uma matriz de cadeias de caracteres que o Outlook reconhece como endereços postais. Da mesma forma, a propriedade **Entities.contacts** retorna uma matriz de objetos **Contact** que o Outlook reconhece como informações de contato. A Tabela 1 lista o tipo de objeto de uma instância de cada entidade compatível.

O exemplo a seguir mostra como recuperar endereços encontrados em uma mensagem.

```js
// Get the address entities from the item.
var entities = Office.context.mailbox.item.getEntities();
// Check to make sure that address entities are present.
if (null != entities && null != entities.addresses && undefined != entities.addresses) {
   //Addresses are present, so use them here.
}

```


## <a name="activating-an-add-in-based-on-the-existence-of-an-entity"></a>Ativar um suplemento com base na existência de uma entidade

Outra maneira de usar entidades conhecidas é fazer com que o Outlook ative o suplemento baseado na existência de um ou mais tipos de entidades no assunto ou no corpo do item exibido no momento. Você pode fazer isso especificando uma regra **ItemHasKnownEntity** no manifesto do suplemento. O tipo simples [EntityType](/javascript/api/outlook/office.mailboxenums.entitytype) representa os diferentes tipos de entidades conhecidas compatíveis com as regras **ItemHasKnownEntity**. Depois de ativar o suplemento, também é possível recuperar as instâncias de tais entidades para seus propósitos, como descrito na seção anterior, [Recuperar entidades no seu suplemento](#retrieving-entities-in-your-add-in).

Você pode opcionalmente aplicar uma expressão regular em uma regra **ItemHasKnownEntity** para filtrar instâncias de uma entidade e fazer com que o Outlook somente ative um suplemento em um subconjunto de instâncias da entidade. Por exemplo, você pode especificar um filtro para a entidade de rua do endereço em uma mensagem que contenha um CEP do Rio de Janeiro que comece com "021". Para aplicar um filtro em instâncias de entidade, use os atributos **RegExFilter** e **FilterName** no elemento `Rule` do tipo [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule).

De forma semelhante às outras regras de ativação, você pode especificar várias regras a fim de formar uma coleção de regras para seu suplemento. O exemplo a seguir aplica-se a uma operação "E" em duas regras: uma regra **ItemIs** e uma regra **ItemHasKnownEntity**. Essa coleção de regras ativa o suplemento sempre que o item atual for uma mensagem e o Outlook reconhecer um endereço no assunto ou no corpo do item.

```XML
<Rule xsi:type="RuleCollection" Mode="And">
   <Rule xsi:type="ItemIs" ItemType="Message" />
   <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
</Rule>
```

<br/>

O exemplo a seguir usa **getEntitiesByType** do item atual para definir uma variável `addresses` nos resultados da coleção de regras anterior.

```js
var addresses = Office.context.mailbox.item.getEntitiesByType(Office.MailboxEnums.EntityType.Address);
```

<br/>

O exemplo de regra **ItemHasKnownEntity** a seguir ativa o suplemento sempre que há uma URL no assunto ou no corpo do item atual, e a URL contém a cadeia de caracteres "youtube" independentemente do uso de maiúsculas e minúsculas na cadeia de caracteres.

```XML
<Rule xsi:type="ItemHasKnownEntity" 
    EntityType="Url" 
    RegExFilter="youtube"
    FilterName="youtube"
    IgnoreCase="true"/>
```

<br/>

O exemplo a seguir usa **getFilteredEntitiesByName(name)** do item atual para definir uma variável `videos` para obter uma matriz de resultados que correspondam a expressões regulares na regra **ItemHasKnownEntity** anterior.

```js
var videos = Office.context.mailbox.item.getFilteredEntitiesByName(youtube);
```


## <a name="tips-for-using-well-known-entities"></a>Dicas para usar entidades conhecidas

Existem alguns fatos e limites de que você deve estar ciente ao usar entidades conhecidas no seu suplemento. Isso se aplica desde que o suplemento seja ativado quando o usuário está lendo um item que contém as correspondências de entidades conhecidas, independentemente de você usar uma regra **ItemHasKnownEntity**:


- Você somente pode extrair cadeias de caracteres que sejam entidades conhecidas se elas estiverem em inglês.
    
- Você pode extrair entidades conhecidas dos primeiros dois mil caracteres no corpo do item, mas não além disso. Esse limite de tamanho ajuda a equilibrar as necessidades de funcionalidade e desempenho, para que o Exchange Server e o Outlook não sejam afetados pela análise e identificação de instâncias de entidades conhecidas em mensagens e compromissos grandes. Observe que esse limite independe de o suplemento especificar uma regra **ItemHasKnownEntity**. Se o suplemento usa uma regra como essa, observe também o limite de processamento de regras no item 2 abaixo para os clientes avançados do Outlook.
    
- Você pode extrair entidades de compromissos que sejam reuniões organizadas por alguém que não seja o proprietário da caixa de correio. Você não pode extrair entidades de itens do calendário que não são reuniões ou reuniões organizadas pelo proprietário da caixa de correio.
    
- Você pode extrair entidades do tipo **MeetingSuggestion** apenas de mensagens, mas não de compromissos.
    
- Você pode extrair URLs que existem explicitamente no corpo do item, mas não URLs que estão inseridas no texto de hiperlink no corpo do item HTML. Em vez disso, considere usar uma regra **ItemHasRegularExpressionMatch** para obter URLs explícitas e inseridas. Especifique **BodyAsHTML** como _PropertyName_ e uma expressão regular que corresponde URLs como _RegExValue_.
    
- Você não pode extrair entidades de itens na pasta Itens Enviados.
    
Além disso, o seguinte se aplica se você usa uma regra [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule), e pode afetar os cenários em que você poderia assumir que seu suplemento seria ativado:

- Ao usar a regra **ItemHasKnownEntity**, assuma que o Outlook corresponderá cadeias de caracteres de entidade somente em inglês, independentemente da localidade padrão especificada no manifesto.
    
- Quando seu suplemento for executado em um cliente avançado do Outlook, assuma que o Outlook aplicará a regra **ItemHasKnownEntity** no primeiro megabyte do corpo do item, e não no restante do corpo acima desse limite.
    
- Você não pode usar uma regra **ItemHasKnownEntity** para ativar um suplemento para itens na pasta Itens Enviados.
    

## <a name="see-also"></a>Confira também

- [Criar suplementos do Outlook para formulários de leitura](read-scenario.md)   
- [Extrair cadeias de caracteres de entidade de um item do Outlook](extract-entity-strings-from-an-item.md)   
- [Regras de ativação para suplementos do Outlook](activation-rules.md)   
- [Usar regras de ativação de expressões regulares para mostrar um suplemento do Outlook](use-regular-expressions-to-show-an-outlook-add-in.md)    
- [Noções básicas sobre permissões de suplemento do Outlook](understanding-outlook-add-in-permissions.md)
    
