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
# <a name="match-strings-in-an-outlook-item-as-well-known-entities"></a><span data-ttu-id="8145b-103">Corresponder cadeias de caracteres em um item do Outlook como entidades conhecidas</span><span class="sxs-lookup"><span data-stu-id="8145b-103">Match strings in an Outlook item as well-known entities</span></span>

<span data-ttu-id="8145b-p101">Antes de enviar um item de mensagem ou de solicitação de reunião, o Exchange Server analisa o conteúdo do item, identifica e apresenta determinadas cadeias de caracteres no assunto e no corpo semelhantes a entidades conhecidas do Exchange, como endereços de email, números de telefone e URLs. As mensagens e solicitações de reunião são fornecidas pelo Exchange Server em uma Caixa de Entrada do Outlook com entidades conhecidas carimbadas.</span><span class="sxs-lookup"><span data-stu-id="8145b-p101">Before sending a message or meeting request item, Exchange Server parses the contents of the item, identifies and stamps certain strings in the subject and body that resemble entities well-known to Exchange, for example, email addresses, phone numbers, and URLs. Messages and meeting requests are delivered by Exchange Server in an Outlook Inbox with well-known entities stamped.</span></span> 

<span data-ttu-id="8145b-p102">Usando a API JavaScript para Office, você pode obter essas cadeias de caracteres que correspondem a entidades conhecidas específicas para processá-las posteriormente. Também pode especificar uma entidade conhecida em uma regra no manifesto do suplemento, para que o Outlook possa ativar o suplemento quando o usuário estiver exibindo um item que contém correspondências para essa entidade. Em seguida, é possível extrair e agir em relação às correspondências da entidade.</span><span class="sxs-lookup"><span data-stu-id="8145b-p102">Using the JavaScript API for Office, you can get these strings that match specific well-known entities for further processing. You can also specify a well-known entity in a rule in the add-in manifest so that Outlook can activate your add-in when the user is viewing an item that contains matches for that entity. You can then extract and take action on matches for the entity.</span></span> 

<span data-ttu-id="8145b-109">Convém ser capaz de identificar ou extrair tais instâncias de uma mensagem ou compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="8145b-109">Being able to identify or extract such instances from a selected message or appointment is convenient.</span></span> <span data-ttu-id="8145b-110">Por exemplo, você pode compilar um serviço de pesquisa invertida de telefones como um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="8145b-110">For example, you can build a reverse phone look-up service as an Outlook add-in.</span></span> <span data-ttu-id="8145b-111">O suplemento pode extrair cadeias de caracteres no corpo ou assunto do item que se parecem com um número de telefone, fazer uma pesquisa invertida e exibir o proprietário registrado de cada número de telefone.</span><span class="sxs-lookup"><span data-stu-id="8145b-111">The add-in can extract strings in the item subject or body that resemble a phone number, do a reverse lookup, and display the registered owner of each phone number.</span></span>

<span data-ttu-id="8145b-112">Este tópico apresenta essas entidades conhecidas, mostra exemplos de regras de ativação baseadas em entidades conhecidas e como extrair correspondências de entidade independentemente de ter usado entidades em regras de ativação.</span><span class="sxs-lookup"><span data-stu-id="8145b-112">This topic introduces these well-known entities, shows examples of activation rules based on well-known entities, and how to extract entity matches independently of having used entities in activation rules.</span></span>


## <a name="support-for-well-known-entities"></a><span data-ttu-id="8145b-113">Suporte para entidades conhecidas</span><span class="sxs-lookup"><span data-stu-id="8145b-113">Support for well-known entities</span></span>

<span data-ttu-id="8145b-p104">O Exchange Server carimba entidades conhecidas em um item de mensagem ou de solicitação de reunião depois que o remetente envia o item e antes de o Exchange entregar o item ao destinatário. Portanto, somente os itens que passaram pelo transporte do Exchange são carimbados, e o Outlook pode ativar suplementos com base nesses carimbos quando o usuário está exibindo esses itens. Do contrário, quando o usuário está redigindo ou visualizando um item que está na pasta Itens Enviados, como o item não passou por transporte, o Outlook não pode ativar suplementos com base em entidades conhecidas.</span><span class="sxs-lookup"><span data-stu-id="8145b-p104">Exchange Server stamps well-known entities in a message or meeting request item after the sender sends the item and before Exchange delivers the item to the recipient. Therefore, only items that have gone through transport in Exchange are stamped, and Outlook can activate add-ins based on these stamps when the user is viewing such items. On the contrary, when the user is composing an item or viewing an item that is in the Sent Items folder, because the item has not gone through transport, Outlook cannot activate add-ins based on well-known entities.</span></span> 

<span data-ttu-id="8145b-p105">Da mesma forma, você não pode extrair entidades conhecidas em itens que estão sendo redigidos ou estão na pasta Itens Enviados, já que esses itens não passaram pelo transporte e não foram carimbados. Para saber mais sobre os tipos de itens que dão suporte à ativação, confira [Regras de ativação para suplementos do Outlook](activation-rules.md).</span><span class="sxs-lookup"><span data-stu-id="8145b-p105">Similarly, you cannot extract well-known entities in items that are being composed or in the Sent Items folder, as these items have not gone through transport and are not stamped. For additional information about the kinds of items that support activation, see [Activation rules for Outlook add-ins](activation-rules.md).</span></span>

<span data-ttu-id="8145b-p106">A tabela a seguir lista as entidades que têm suporte e são reconhecidas pelo Exchange Server e pelo Outlook (por isso chamadas "entidades conhecidas") e o tipo de objeto de uma instância de cada entidade. O reconhecimento de linguagem natural de uma cadeia de caracteres como uma dessas entidades baseia-se em um modelo de aprendizagem que foi treinado com grande quantidade de dados. Portanto, o reconhecimento é não determinístico. Confira [Dicas para usar entidades conhecidas](#tips-for-using-well-known-entities) a fim de saber mais sobre condições de reconhecimento.</span><span class="sxs-lookup"><span data-stu-id="8145b-p106">The following table lists the entities that Exchange Server and Outlook support and recognize (hence the name "well-known entities"), and the object type of an instance of each entity. The natural language recognition of a string as one of these entities is based on a learning model that has been trained on a large amount of data. Therefore, the recognition is non-deterministic. See [Tips for using well-known entities](#tips-for-using-well-known-entities) for more information about conditions for recognition.</span></span>

<span data-ttu-id="8145b-123">**Tabela 1. Entidades compatíveis e os respectivos tipos**</span><span class="sxs-lookup"><span data-stu-id="8145b-123">**Table 1. Supported entities and their types**</span></span>

|<span data-ttu-id="8145b-124">Tipo de entidade</span><span class="sxs-lookup"><span data-stu-id="8145b-124">Entity type</span></span>|<span data-ttu-id="8145b-125">Condições de reconhecimento</span><span class="sxs-lookup"><span data-stu-id="8145b-125">Conditions for recognition</span></span>|<span data-ttu-id="8145b-126">Tipo de objeto</span><span class="sxs-lookup"><span data-stu-id="8145b-126">Object type</span></span>|
|:-----|:-----|:-----|
|<span data-ttu-id="8145b-127">**Endereço**</span><span class="sxs-lookup"><span data-stu-id="8145b-127">**Address**</span></span>|<span data-ttu-id="8145b-p107">Endereços nos Estados Unidos. Por exemplo: 1234 Main Street, Redmond, WA 07722. Normalmente, para um endereço ser reconhecido, ele deve seguir a estrutura de um endereço postal dos Estados Unidos, com a maioria dos elementos de nome da rua, número, cidade, estado e CEP. O endereço pode ser especificado em uma ou várias linhas.</span><span class="sxs-lookup"><span data-stu-id="8145b-p107">United States street addresses; for example: 1234 Main Street, Redmond, WA 07722. Generally, for an address to be recognized, it should follow the structure of a United States postal address, with most of the elements of a street number, street name, city, state, and zip code present. The address can be specified in one or multiple lines.</span></span>|<span data-ttu-id="8145b-131">Objeto JavaScript **String**</span><span class="sxs-lookup"><span data-stu-id="8145b-131">JavaScript **String** object</span></span>|
|<span data-ttu-id="8145b-132">**Contato**</span><span class="sxs-lookup"><span data-stu-id="8145b-132">**Contact**</span></span>|<span data-ttu-id="8145b-133">Uma referência às informações de uma pessoa assim reconhecida em sua língua materna.</span><span class="sxs-lookup"><span data-stu-id="8145b-133">A reference to a person's information as recognized in natural language.</span></span> <span data-ttu-id="8145b-134">O reconhecimento de um contato depende do contexto.</span><span class="sxs-lookup"><span data-stu-id="8145b-134">The recognition of a contact depends on the context.</span></span> <span data-ttu-id="8145b-135">Por exemplo, uma assinatura no final de uma mensagem ou o nome da pessoa que aparece perto de algumas das seguintes informações: número de telefone, endereço, endereço de e-mail e URL.</span><span class="sxs-lookup"><span data-stu-id="8145b-135">For example, a signature at the end of a message, or a person's name appearing in the vicinity of some of the following information: a phone number, address, email address, and URL.</span></span>|<span data-ttu-id="8145b-136">Objeto [Contact](/javascript/api/outlook/office.contact)</span><span class="sxs-lookup"><span data-stu-id="8145b-136">[Contact](/javascript/api/outlook/office.contact) object</span></span>|
|<span data-ttu-id="8145b-137">**EmailAddress**</span><span class="sxs-lookup"><span data-stu-id="8145b-137">**EmailAddress**</span></span>|<span data-ttu-id="8145b-138">Endereços de email SMTP.</span><span class="sxs-lookup"><span data-stu-id="8145b-138">SMTP email addresses.</span></span>|<span data-ttu-id="8145b-139">Objeto JavaScript **String**</span><span class="sxs-lookup"><span data-stu-id="8145b-139">JavaScript **String** object</span></span>|
|<span data-ttu-id="8145b-140">**MeetingSuggestion**</span><span class="sxs-lookup"><span data-stu-id="8145b-140">**MeetingSuggestion**</span></span>|<span data-ttu-id="8145b-p109">Uma referência a uma reunião ou a um evento. Por exemplo, o Exchange 2013 reconheceria o seguinte texto como uma sugestão de reunião:  _Vamos marcar um almoço amanhã._</span><span class="sxs-lookup"><span data-stu-id="8145b-p109">A reference to an event or meeting. For example, Exchange 2013 would recognize the following text as a meeting suggestion:  _Let's meet tomorrow for lunch._</span></span>|<span data-ttu-id="8145b-143">Objeto [MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)</span><span class="sxs-lookup"><span data-stu-id="8145b-143">[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion) object</span></span>|
|<span data-ttu-id="8145b-144">**PhoneNumber**</span><span class="sxs-lookup"><span data-stu-id="8145b-144">**PhoneNumber**</span></span>|<span data-ttu-id="8145b-145">Números de telefone dos Estados Unidos. Por exemplo:  _(235) 555-0110_</span><span class="sxs-lookup"><span data-stu-id="8145b-145">United States telephone numbers; for example:  _(235) 555-0110_</span></span>|<span data-ttu-id="8145b-146">Objeto [PhoneNumber](/javascript/api/outlook/office.phonenumber)</span><span class="sxs-lookup"><span data-stu-id="8145b-146">[PhoneNumber](/javascript/api/outlook/office.phonenumber) object</span></span>|
|<span data-ttu-id="8145b-147">**TaskSuggestion**</span><span class="sxs-lookup"><span data-stu-id="8145b-147">**TaskSuggestion**</span></span>|<span data-ttu-id="8145b-p110">Frases acionáveis em um email. Por exemplo:  _Atualize a planilha._</span><span class="sxs-lookup"><span data-stu-id="8145b-p110">Actionable sentences in an email. For example:  _Please update the spreadsheet._</span></span>|<span data-ttu-id="8145b-150">Objeto [TaskSuggestion](/javascript/api/outlook/office.tasksuggestion)</span><span class="sxs-lookup"><span data-stu-id="8145b-150">[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion) object</span></span>|
|<span data-ttu-id="8145b-151">**Url**</span><span class="sxs-lookup"><span data-stu-id="8145b-151">**Url**</span></span>|<span data-ttu-id="8145b-152">Um endereço Web que especifica explicitamente o local de rede e o identificador de um recurso da Web.</span><span class="sxs-lookup"><span data-stu-id="8145b-152">A web address that explicitly specifies the network location and identifier for a web resource.</span></span> <span data-ttu-id="8145b-153">O Exchange Server não requer o protocolo de acesso no endereço Web e não reconhece URLs que são inseridas no texto do link como instâncias da entidade **Url**.</span><span class="sxs-lookup"><span data-stu-id="8145b-153">Exchange Server does not require the access protocol in the web address, and does not recognize URLs that are embedded in link text as instances of the **Url** entity.</span></span> <span data-ttu-id="8145b-154">O Exchange Server pode corresponder aos seguintes exemplos `www.youtube.com/user/officevideos` :`https://www.youtube.com/user/officevideos`</span><span class="sxs-lookup"><span data-stu-id="8145b-154">Exchange Server can match the following examples: `www.youtube.com/user/officevideos` `https://www.youtube.com/user/officevideos`</span></span> |<span data-ttu-id="8145b-155">Objeto JavaScript **String**</span><span class="sxs-lookup"><span data-stu-id="8145b-155">JavaScript **String** object</span></span>|

<br/>

<span data-ttu-id="8145b-p112">A figura a seguir descreve como o Exchange Server e o Outlook dão suporte a entidades conhecidas para suplementos e o que estes podem fazer com entidades conhecidas. Confira [Recuperar entidades em seu suplemento](#retrieving-entities-in-your-add-in) e [Ativar um suplemento com base na existência de uma entidade](#activating-an-add-in-based-on-the-existence-of-an-entity) para obter mais detalhes sobre como usar essas entidades.</span><span class="sxs-lookup"><span data-stu-id="8145b-p112">The following figure describes how Exchange Server and Outlook support well-known entities for add-ins, and what add-ins can do with well-known entities. See [Retrieving entities in your add-in](#retrieving-entities-in-your-add-in) and [Activating an add-in based on the existence of an entity](#activating-an-add-in-based-on-the-existence-of-an-entity) for more details on how to use these entities.</span></span>

<span data-ttu-id="8145b-158">**Como o Exchange Server, o Outlook e os suplementos dão suporte a entidades conhecidas**</span><span class="sxs-lookup"><span data-stu-id="8145b-158">**How Exchange Server, Outlook, and add-ins support well-known entities**</span></span>

![Support and use of well-known entities in mail app](../images/well-known-entities-info.png)


## <a name="permissions-to-extract-entities"></a><span data-ttu-id="8145b-160">Permissões para extrair entidades</span><span class="sxs-lookup"><span data-stu-id="8145b-160">Permissions to extract entities</span></span>

<span data-ttu-id="8145b-161">Para extrair entidades no seu código JavaScript ou fazer com que seu suplemento seja ativado com base na existência de determinadas entidades conhecidas, verifique se você solicitou as permissões apropriadas no manifesto do suplemento.</span><span class="sxs-lookup"><span data-stu-id="8145b-161">To extract entities in your JavaScript code or to have your add-in activated based on the existence of certain well-known entities, make sure you have requested the appropriate permissions in the add-in manifest.</span></span>

<span data-ttu-id="8145b-162">A especificação da permissão restrita padrão permite que o suplemento extraia as entidades **Address**, **MeetingSuggestion** ou **TaskSuggestion**.</span><span class="sxs-lookup"><span data-stu-id="8145b-162">Specifying the default restricted permission allows your add-in to extract the **Address**, **MeetingSuggestion**, or **TaskSuggestion** entity.</span></span> <span data-ttu-id="8145b-163">Para extrair as outras entidades, especifique as permissões de leitura de item, leitura/gravação de item ou leitura/gravação de caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="8145b-163">To extract any of the other entities, specify read item, read/write item, or read/write mailbox permission.</span></span> <span data-ttu-id="8145b-164">Para fazer isso no manifesto, use o elemento [Permissions](../reference/manifest/permissions.md) e especifique a permissão apropriada &mdash; **Restricted**, **ReadItem**, **ReadWriteItem** ou **ReadWriteMailbox** &mdash; como no exemplo abaixo:</span><span class="sxs-lookup"><span data-stu-id="8145b-164">To do that in the manifest, use the [Permissions](../reference/manifest/permissions.md) element and specify the appropriate permission&mdash;**Restricted**, **ReadItem**, **ReadWriteItem**, or **ReadWriteMailbox**&mdash;as in the following example:</span></span>

```xml
<Permissions>ReadItem</Permissions>
```


## <a name="retrieving-entities-in-your-add-in"></a><span data-ttu-id="8145b-165">Recuperar entidades no seu suplemento</span><span class="sxs-lookup"><span data-stu-id="8145b-165">Retrieving entities in your add-in</span></span>

<span data-ttu-id="8145b-p114">Desde que o assunto ou corpo do item que está sendo visualizado pelo usuário contenha cadeias de caracteres que o Exchange e o Outlook possam reconhecer como entidades conhecidas, essas instâncias estarão disponíveis para suplementos. Elas estarão disponíveis mesmo que um suplemento não seja ativado com base em entidades conhecidas. Com a permissão apropriada, você pode usar os métodos **getEntities** ou **getEntitiesByType** para recuperar entidades conhecidas que estejam presentes na mensagem ou compromisso atual.</span><span class="sxs-lookup"><span data-stu-id="8145b-p114">As long as the subject or body of the item that is being viewed by the user contains strings that Exchange and Outlook can recognize as well-known entities, these instances are available to add-ins. They are available even if an add-in is not activated based on well-known entities. With the appropriate permission, you can use the **getEntities** or **getEntitiesByType** method to retrieve well-known entities that are present in the current message or appointment.</span></span>

<span data-ttu-id="8145b-168">O método **getEntities** retorna uma matriz de objetos [Entities](/javascript/api/outlook/office.entities) que contém todas as entidades conhecidas no item.</span><span class="sxs-lookup"><span data-stu-id="8145b-168">The **getEntities** method returns an array of [Entities](/javascript/api/outlook/office.entities) objects that contains all the well-known entities in the item.</span></span>

<span data-ttu-id="8145b-169">Se você estiver interessado em um determinado tipo de entidades, use o método **getEntitiesByType**, que retorna uma matriz somente com as entidades desejadas.</span><span class="sxs-lookup"><span data-stu-id="8145b-169">If you're interested in a particular type of entities, use the **getEntitiesByType** method which returns an array of only the entities you want.</span></span> <span data-ttu-id="8145b-170">A enumeração [EntityType](/javascript/api/outlook/office.mailboxenums.entitytype) representa todos os tipos de entidades conhecidas que você pode extrair.</span><span class="sxs-lookup"><span data-stu-id="8145b-170">The [EntityType](/javascript/api/outlook/office.mailboxenums.entitytype) enumeration represents all the types of well-known entities you can extract.</span></span>

<span data-ttu-id="8145b-171">Após chamar **getEntities**, você pode usar a propriedade correspondente do objeto **Entities** para obter uma matriz de instâncias de um tipo de entidade.</span><span class="sxs-lookup"><span data-stu-id="8145b-171">After calling **getEntities**, you can then use the corresponding property of the **Entities** object to obtain an array of instances of a type of entity.</span></span> <span data-ttu-id="8145b-172">Dependendo do tipo de entidade, as instâncias na matriz podem ser apenas cadeias de caracteres ou podem mapear para objetos específicos.</span><span class="sxs-lookup"><span data-stu-id="8145b-172">Depending on the type of entity, the instances in the array can be just strings, or can map to specific objects.</span></span> 

<span data-ttu-id="8145b-173">Como o exemplo mostrado na figura anterior, acesse a matriz retornada por `getEntities().addresses[]` para obter endereços no item.</span><span class="sxs-lookup"><span data-stu-id="8145b-173">As an example seen in the earlier figure, to get addresses in the item, access the array returned by `getEntities().addresses[]`.</span></span> <span data-ttu-id="8145b-174">A propriedade **Entities.addresses** retorna uma matriz de cadeias de caracteres que o Outlook reconhece como endereços postais.</span><span class="sxs-lookup"><span data-stu-id="8145b-174">The **Entities.addresses** property returns an array of strings that Outlook recognizes as postal addresses.</span></span> <span data-ttu-id="8145b-175">Da mesma forma, a propriedade **Entities.contacts** retorna uma matriz de objetos **Contact** que o Outlook reconhece como informações de contato.</span><span class="sxs-lookup"><span data-stu-id="8145b-175">Similarly, the **Entities.contacts** property returns an array of **Contact** objects that Outlook recognizes as contact information.</span></span> <span data-ttu-id="8145b-176">A Tabela 1 lista o tipo de objeto de uma instância de cada entidade compatível.</span><span class="sxs-lookup"><span data-stu-id="8145b-176">Tables 1 lists the object type of an instance of each supported entity.</span></span>

<span data-ttu-id="8145b-177">O exemplo a seguir mostra como recuperar endereços encontrados em uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="8145b-177">The following example shows how to retrieve any addresses found in a message.</span></span>

```js
// Get the address entities from the item.
var entities = Office.context.mailbox.item.getEntities();
// Check to make sure that address entities are present.
if (null != entities && null != entities.addresses && undefined != entities.addresses) {
   //Addresses are present, so use them here.
}

```


## <a name="activating-an-add-in-based-on-the-existence-of-an-entity"></a><span data-ttu-id="8145b-178">Ativar um suplemento com base na existência de uma entidade</span><span class="sxs-lookup"><span data-stu-id="8145b-178">Activating an add-in based on the existence of an entity</span></span>

<span data-ttu-id="8145b-p118">Outra maneira de usar entidades conhecidas é fazer com que o Outlook ative o suplemento baseado na existência de um ou mais tipos de entidades no assunto ou no corpo do item exibido no momento. Você pode fazer isso especificando uma regra **ItemHasKnownEntity** no manifesto do suplemento. O tipo simples [EntityType](/javascript/api/outlook/office.mailboxenums.entitytype) representa os diferentes tipos de entidades conhecidas compatíveis com as regras **ItemHasKnownEntity**. Depois de ativar o suplemento, também é possível recuperar as instâncias de tais entidades para seus propósitos, como descrito na seção anterior, [Recuperar entidades no seu suplemento](#retrieving-entities-in-your-add-in).</span><span class="sxs-lookup"><span data-stu-id="8145b-p118">Another way to use well-known entities is to have Outlook activate your add-in based on the existence of one or more types of entities in the subject or body of the currently viewed item. You can do so by specifying an **ItemHasKnownEntity** rule in the add-in manifest. The [EntityType](/javascript/api/outlook/office.mailboxenums.entitytype) simple type represents the different types of well-known entities supported by **ItemHasKnownEntity** rules. After your add-in is activated, you can also retrieve the instances of such entities for your purposes, as described in the previous section [Retrieving entities in your add-in](#retrieving-entities-in-your-add-in).</span></span>

<span data-ttu-id="8145b-p119">Você pode opcionalmente aplicar uma expressão regular em uma regra **ItemHasKnownEntity** para filtrar instâncias de uma entidade e fazer com que o Outlook somente ative um suplemento em um subconjunto de instâncias da entidade. Por exemplo, você pode especificar um filtro para a entidade de rua do endereço em uma mensagem que contenha um CEP do Rio de Janeiro que comece com "021". Para aplicar um filtro em instâncias de entidade, use os atributos **RegExFilter** e **FilterName** no elemento `Rule` do tipo [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule).</span><span class="sxs-lookup"><span data-stu-id="8145b-p119">You can optionally apply a regular expression in an **ItemHasKnownEntity** rule, so as to further filter instances of an entity and have Outlook activate an add-in only on a subset of the instances of the entity. For example, you can specify a filter for the street address entity in a message that contains a Washington state zip code beginning with "98". To apply a filter on the entity instances, use the **RegExFilter** and **FilterName** attributes in the `Rule` element of the [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) type.</span></span>

<span data-ttu-id="8145b-p120">De forma semelhante às outras regras de ativação, você pode especificar várias regras a fim de formar uma coleção de regras para seu suplemento. O exemplo a seguir aplica-se a uma operação "E" em duas regras: uma regra **ItemIs** e uma regra **ItemHasKnownEntity**. Essa coleção de regras ativa o suplemento sempre que o item atual for uma mensagem e o Outlook reconhecer um endereço no assunto ou no corpo do item.</span><span class="sxs-lookup"><span data-stu-id="8145b-p120">Similar to other activation rules, you can specify multiple rules to form a rule collection for your add-in. The following example applies an "AND" operation on 2 rules: an **ItemIs** rule and an **ItemHasKnownEntity** rule. This rule collection activates the add-in whenever the current item is a message and Outlook recognizes an address in the subject or body of that item.</span></span>

```XML
<Rule xsi:type="RuleCollection" Mode="And">
   <Rule xsi:type="ItemIs" ItemType="Message" />
   <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
</Rule>
```

<br/>

<span data-ttu-id="8145b-189">O exemplo a seguir usa **getEntitiesByType** do item atual para definir uma variável `addresses` nos resultados da coleção de regras anterior.</span><span class="sxs-lookup"><span data-stu-id="8145b-189">The following example uses **getEntitiesByType** of the current item to set a variable `addresses` to the results of the preceding rule collection.</span></span>

```js
var addresses = Office.context.mailbox.item.getEntitiesByType(Office.MailboxEnums.EntityType.Address);
```

<br/>

<span data-ttu-id="8145b-190">O exemplo de regra **ItemHasKnownEntity** a seguir ativa o suplemento sempre que há uma URL no assunto ou no corpo do item atual, e a URL contém a cadeia de caracteres "youtube" independentemente do uso de maiúsculas e minúsculas na cadeia de caracteres.</span><span class="sxs-lookup"><span data-stu-id="8145b-190">The following **ItemHasKnownEntity** rule example activates the add-in whenever there is a URL in the subject or body of the current item, and the URL contains the string "youtube", regardless of the case of the string.</span></span>

```XML
<Rule xsi:type="ItemHasKnownEntity" 
    EntityType="Url" 
    RegExFilter="youtube"
    FilterName="youtube"
    IgnoreCase="true"/>
```

<br/>

<span data-ttu-id="8145b-191">O exemplo a seguir usa **getFilteredEntitiesByName(name)** do item atual para definir uma variável `videos` para obter uma matriz de resultados que correspondam a expressões regulares na regra **ItemHasKnownEntity** anterior.</span><span class="sxs-lookup"><span data-stu-id="8145b-191">The following example uses **getFilteredEntitiesByName(name)** of the current item to set a variable `videos` to get an array of results that match the regular expression in the preceding **ItemHasKnownEntity** rule.</span></span>

```js
var videos = Office.context.mailbox.item.getFilteredEntitiesByName(youtube);
```


## <a name="tips-for-using-well-known-entities"></a><span data-ttu-id="8145b-192">Dicas para usar entidades conhecidas</span><span class="sxs-lookup"><span data-stu-id="8145b-192">Tips for using well-known entities</span></span>

<span data-ttu-id="8145b-p121">Existem alguns fatos e limites de que você deve estar ciente ao usar entidades conhecidas no seu suplemento. Isso se aplica desde que o suplemento seja ativado quando o usuário está lendo um item que contém as correspondências de entidades conhecidas, independentemente de você usar uma regra **ItemHasKnownEntity**:</span><span class="sxs-lookup"><span data-stu-id="8145b-p121">There are a few facts and limits you should be aware of if you use well-known entities in your add-in. The following applies as long as your add-in is activated when the user is reading an item which contains matches of well-known entities, regardless of whether you use an **ItemHasKnownEntity** rule:</span></span>


- <span data-ttu-id="8145b-195">Você somente pode extrair cadeias de caracteres que sejam entidades conhecidas se elas estiverem em inglês.</span><span class="sxs-lookup"><span data-stu-id="8145b-195">You can extract strings that are well-known entities only if the strings are in English.</span></span>
    
- <span data-ttu-id="8145b-196">Você pode extrair entidades conhecidas dos primeiros dois mil caracteres no corpo do item, mas não além disso.</span><span class="sxs-lookup"><span data-stu-id="8145b-196">You can extract well-known entities from the first 2,000 characters in the item body, but not beyond that limit.</span></span> <span data-ttu-id="8145b-197">Esse limite de tamanho ajuda a equilibrar as necessidades de funcionalidade e desempenho, para que o Exchange Server e o Outlook não sejam afetados pela análise e identificação de instâncias de entidades conhecidas em mensagens e compromissos grandes.</span><span class="sxs-lookup"><span data-stu-id="8145b-197">This size limit helps balance the need for functionality and performance, so that Exchange Server and Outlook are not bogged down by parsing and identifying instances of well-known entities in large messages and appointments.</span></span> <span data-ttu-id="8145b-198">Observe que esse limite independe de o suplemento especificar uma regra **ItemHasKnownEntity**.</span><span class="sxs-lookup"><span data-stu-id="8145b-198">Note that this limit is independent of whether the add-in specifies an **ItemHasKnownEntity** rule.</span></span> <span data-ttu-id="8145b-199">Se o suplemento usa uma regra como essa, observe também o limite de processamento de regras no item 2 abaixo para os clientes avançados do Outlook.</span><span class="sxs-lookup"><span data-stu-id="8145b-199">If the add-in does use such a rule, note also the rule processing limit in item 2 below for the Outlook rich clients.</span></span>
    
- <span data-ttu-id="8145b-p123">Você pode extrair entidades de compromissos que sejam reuniões organizadas por alguém que não seja o proprietário da caixa de correio. Você não pode extrair entidades de itens do calendário que não são reuniões ou reuniões organizadas pelo proprietário da caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="8145b-p123">You can extract entities from appointments that are meetings organized by someone other than the mailbox owner. You cannot extract entities from calendar items that are not meetings, or meetings organized by the mailbox owner.</span></span>
    
- <span data-ttu-id="8145b-202">Você pode extrair entidades do tipo **MeetingSuggestion** apenas de mensagens, mas não de compromissos.</span><span class="sxs-lookup"><span data-stu-id="8145b-202">You can extract entities of the **MeetingSuggestion** type from only messages but not appointments.</span></span>
    
- <span data-ttu-id="8145b-p124">Você pode extrair URLs que existem explicitamente no corpo do item, mas não URLs que estão inseridas no texto de hiperlink no corpo do item HTML. Em vez disso, considere usar uma regra **ItemHasRegularExpressionMatch** para obter URLs explícitas e inseridas. Especifique **BodyAsHTML** como _PropertyName_ e uma expressão regular que corresponde URLs como _RegExValue_.</span><span class="sxs-lookup"><span data-stu-id="8145b-p124">You can extract URLs that exist explicitly in the item body, but not URLs that are embedded in hyperlinked text in HTML item body. Consider using an **ItemHasRegularExpressionMatch** rule instead to get both explicit and embedded URLs. Specify **BodyAsHTML** as the _PropertyName_, and a regular expression that matches URLs as the  _RegExValue_.</span></span>
    
- <span data-ttu-id="8145b-206">Você não pode extrair entidades de itens na pasta Itens Enviados.</span><span class="sxs-lookup"><span data-stu-id="8145b-206">You cannot extract entities from items in the Sent Items folder.</span></span>
    
<span data-ttu-id="8145b-207">Além disso, o seguinte se aplica se você usa uma regra [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule), e pode afetar os cenários em que você poderia assumir que seu suplemento seria ativado:</span><span class="sxs-lookup"><span data-stu-id="8145b-207">In addition, the following applies if you use an [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) rule, and may affect the scenarios where you'd otherwise expect your add-in to be activated:</span></span>

- <span data-ttu-id="8145b-208">Ao usar a regra **ItemHasKnownEntity**, assuma que o Outlook corresponderá cadeias de caracteres de entidade somente em inglês, independentemente da localidade padrão especificada no manifesto.</span><span class="sxs-lookup"><span data-stu-id="8145b-208">When using the **ItemHasKnownEntity** rule, expect Outlook to match entity strings in only English regardless of the default locale specified in the manifest.</span></span>
    
- <span data-ttu-id="8145b-209">Quando seu suplemento for executado em um cliente avançado do Outlook, assuma que o Outlook aplicará a regra **ItemHasKnownEntity** no primeiro megabyte do corpo do item, e não no restante do corpo acima desse limite.</span><span class="sxs-lookup"><span data-stu-id="8145b-209">When your add-in is running on an Outlook rich client, expect Outlook to apply the **ItemHasKnownEntity** rule to the first megabyte of the item body and not to the rest of the body over that limit.</span></span>
    
- <span data-ttu-id="8145b-210">Você não pode usar uma regra **ItemHasKnownEntity** para ativar um suplemento para itens na pasta Itens Enviados.</span><span class="sxs-lookup"><span data-stu-id="8145b-210">You cannot use an **ItemHasKnownEntity** rule to activate an add-in for items in the Sent Items folder.</span></span>
    

## <a name="see-also"></a><span data-ttu-id="8145b-211">Confira também</span><span class="sxs-lookup"><span data-stu-id="8145b-211">See also</span></span>

- [<span data-ttu-id="8145b-212">Criar suplementos do Outlook para formulários de leitura</span><span class="sxs-lookup"><span data-stu-id="8145b-212">Create Outlook add-ins for read forms</span></span>](read-scenario.md)   
- [<span data-ttu-id="8145b-213">Extrair cadeias de caracteres de entidade de um item do Outlook</span><span class="sxs-lookup"><span data-stu-id="8145b-213">Extract entity strings from an Outlook item</span></span>](extract-entity-strings-from-an-item.md)   
- [<span data-ttu-id="8145b-214">Regras de ativação para suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="8145b-214">Activation rules for Outlook add-ins</span></span>](activation-rules.md)   
- [<span data-ttu-id="8145b-215">Usar regras de ativação de expressões regulares para mostrar um suplemento do Outlook</span><span class="sxs-lookup"><span data-stu-id="8145b-215">Use regular expression activation rules to show an Outlook add-in</span></span>](use-regular-expressions-to-show-an-outlook-add-in.md)    
- [<span data-ttu-id="8145b-216">Noções básicas sobre permissões de suplemento do Outlook</span><span class="sxs-lookup"><span data-stu-id="8145b-216">Understanding Outlook add-in permissions</span></span>](understanding-outlook-add-in-permissions.md)
    