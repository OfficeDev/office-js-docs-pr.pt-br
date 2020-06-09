---
title: Suplementos do Outlook para o Outlook Mobile
description: Os suplementos do Outlook Mobile têm suporte em todas as contas do Office 365 Comercial, Outlook.com e, em breve, haverá suporte para contas do Gmail.
ms.date: 05/27/2020
localization_priority: Normal
ms.openlocfilehash: 9a7345840d5a26b27f824470efd58d846d0aab11
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44606438"
---
# <a name="add-ins-for-outlook-mobile"></a><span data-ttu-id="0f96f-103">Suplementos do Outlook Mobile</span><span class="sxs-lookup"><span data-stu-id="0f96f-103">Add-ins for Outlook Mobile</span></span>

<span data-ttu-id="0f96f-p101">Os suplementos agora funcionam no Outlook Mobile, usando as mesmas APIs disponíveis para outros pontos de extremidade do Outlook. Se você já tiver criado um suplemento para Outlook, é fácil fazê-lo funcionar no Outlook Mobile.</span><span class="sxs-lookup"><span data-stu-id="0f96f-p101">Add-ins now work on Outlook Mobile, using the same APIs available for other Outlook endpoints. If you've built an add-in for Outlook already, it's easy to get it working on Outlook Mobile.</span></span>

<span data-ttu-id="0f96f-106">Os suplementos do Outlook Mobile têm suporte em todas as contas do Office 365 Comercial, Outlook.com e, em breve, haverá suporte para contas do Gmail.</span><span class="sxs-lookup"><span data-stu-id="0f96f-106">Outlook mobile add-ins are supported on all Office 365 Commercial accounts, Outlook.com accounts, and support is coming soon to Gmail accounts.</span></span>

<span data-ttu-id="0f96f-107">**Um painel de tarefas de exemplo no Outlook no iOS**</span><span class="sxs-lookup"><span data-stu-id="0f96f-107">**An example task pane in Outlook on iOS**</span></span>

![Uma captura de tela do painel de tarefas no Outlook no iOS](../images/outlook-mobile-addin-taskpane.png)

<br/>

<span data-ttu-id="0f96f-109">**Um painel de tarefas de exemplo no Outlook no Android**</span><span class="sxs-lookup"><span data-stu-id="0f96f-109">**An example task pane in Outlook on Android**</span></span>

![Uma captura de tela do painel de tarefas no Outlook no Android](../images/outlook-mobile-addin-taskpane-android.png)

> [!IMPORTANT]
> <span data-ttu-id="0f96f-111">Os suplementos não funcionam na versão moderna do Outlook em um navegador móvel.</span><span class="sxs-lookup"><span data-stu-id="0f96f-111">Add-ins don't work in the modern version of Outlook in a mobile browser.</span></span> <span data-ttu-id="0f96f-112">Para obter mais informações, consulte [Outlook em seu navegador móvel está sendo atualizado](https://techcommunity.microsoft.com/t5/outlook-blog/outlook-on-your-mobile-browser-is-being-upgraded/ba-p/1125816).</span><span class="sxs-lookup"><span data-stu-id="0f96f-112">For more information, see [Outlook on your mobile browser is being upgraded](https://techcommunity.microsoft.com/t5/outlook-blog/outlook-on-your-mobile-browser-is-being-upgraded/ba-p/1125816).</span></span>

## <a name="whats-different-on-mobile"></a><span data-ttu-id="0f96f-113">Qual é a diferença no celular?</span><span class="sxs-lookup"><span data-stu-id="0f96f-113">What's different on mobile?</span></span>

- <span data-ttu-id="0f96f-p103">O tamanho pequeno e as rápidas interações tornam o projeto para celular um desafio. Para garantir experiências de qualidade para nossos clientes, estamos definindo critérios rígidos de validação que devem ser cumpridos por um suplemento que declara suporte a celular de forma a ser aprovado na AppSource.</span><span class="sxs-lookup"><span data-stu-id="0f96f-p103">The small size and quick interactions make designing for mobile a challenge. To ensure quality experiences for our customers, we are setting strict validation criteria that must be met by an add-in declaring mobile support, in order to be approved in AppSource.</span></span>
    - <span data-ttu-id="0f96f-116">O suplemento **DEVE** cumprir as [diretrizes de interface do usuário](outlook-addin-design.md).</span><span class="sxs-lookup"><span data-stu-id="0f96f-116">The add-in **MUST** adhere to the [UI guidelines](outlook-addin-design.md).</span></span>
    - <span data-ttu-id="0f96f-117">O cenário do suplemento **DEVE** [fazer sentido no mobile](#what-makes-a-good-scenario-for-mobile-add-ins).</span><span class="sxs-lookup"><span data-stu-id="0f96f-117">The scenario for the add-in **MUST** [make sense on mobile](#what-makes-a-good-scenario-for-mobile-add-ins).</span></span>

- <span data-ttu-id="0f96f-118">Em geral, só há suporte para o modo de leitura de mensagens no momento.</span><span class="sxs-lookup"><span data-stu-id="0f96f-118">In general, only Message Read mode is supported at this time.</span></span> <span data-ttu-id="0f96f-119">Isso significa que `MobileMessageReadCommandSurface` é o único [ExtensionPoint](../reference/manifest/extensionpoint.md#mobilemessagereadcommandsurface) que você deve declarar na seção móvel do seu manifesto.</span><span class="sxs-lookup"><span data-stu-id="0f96f-119">That means `MobileMessageReadCommandSurface` is the only [ExtensionPoint](../reference/manifest/extensionpoint.md#mobilemessagereadcommandsurface) you should declare in the mobile section of your manifest.</span></span> <span data-ttu-id="0f96f-120">No entanto, o modo organizador de compromissos tem suporte para suplementos integrados de provedor de reunião online que, em vez disso, declare o [ponto de extensão MobileOnlineMeetingCommandSurface](../reference/manifest/extensionpoint.md#mobileonlinemeetingcommandsurface-preview).</span><span class="sxs-lookup"><span data-stu-id="0f96f-120">However, Appointment Organizer mode is supported for online meeting provider integrated add-ins which instead declare the [MobileOnlineMeetingCommandSurface extension point](../reference/manifest/extensionpoint.md#mobileonlinemeetingcommandsurface-preview).</span></span> <span data-ttu-id="0f96f-121">Confira o artigo [criar um suplemento do Outlook Mobile para um provedor de reunião online](online-meeting.md) para saber mais sobre esse cenário.</span><span class="sxs-lookup"><span data-stu-id="0f96f-121">See the [Create an Outlook mobile add-in for an online-meeting provider](online-meeting.md) article for more about this scenario.</span></span>

- <span data-ttu-id="0f96f-p105">A API [makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) não é suportada no celular, já que o aplicativo móvel usa APIs REST para se comunicar com o servidor. Se seu back-end do aplicativo precisa se conectar ao servidor do Exchange, é possível usar o token de retorno de chamada para fazer chamadas de API REST. Para obter detalhes, consulte [Usar APIs REST do Outlook de um suplemento do Outlook](use-rest-api.md).</span><span class="sxs-lookup"><span data-stu-id="0f96f-p105">The [makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) API is not supported on mobile since the mobile app uses REST APIs to communicate with the server. If your app backend needs to connect to the Exchange server, you can use the callback token to make REST API calls. For details, see [Use the Outlook REST APIs from an Outlook add-in](use-rest-api.md).</span></span>

- <span data-ttu-id="0f96f-125">Quando você envia o suplemento para a loja com [MobileFormFactor](../reference/manifest/mobileformfactor.md) no manifesto, precisará concordar com nosso adendo de suplementos no iOS e precisará enviar sua ID de desenvolvedor Apple para verificação.</span><span class="sxs-lookup"><span data-stu-id="0f96f-125">When you submit your add-in to the store with [MobileFormFactor](../reference/manifest/mobileformfactor.md) in the manifest, you'll need to agree to our developer addendum for add-ins on iOS, and you must submit your Apple Developer ID for verification.</span></span>

- <span data-ttu-id="0f96f-126">Por fim, seu manifesto precisará declarar `MobileFormFactor` e ter os tipos corretos de [controles](../reference/manifest/control.md) e [tamanhos de ícone](../reference/manifest/icon.md) incluídos.</span><span class="sxs-lookup"><span data-stu-id="0f96f-126">Finally, your manifest will need to declare `MobileFormFactor`, and have the correct types of [controls](../reference/manifest/control.md) and [icon sizes](../reference/manifest/icon.md) included.</span></span>

## <a name="what-makes-a-good-scenario-for-mobile-add-ins"></a><span data-ttu-id="0f96f-127">O que forma um bom cenário para suplementos móveis?</span><span class="sxs-lookup"><span data-stu-id="0f96f-127">What makes a good scenario for mobile add-ins?</span></span>

<span data-ttu-id="0f96f-p106">Lembre-se de que o tamanho médio da sessão Outlook em um telefone é bem menor do que em um PC. Isso significa que seu suplemento deve ser rápido e o cenário deve permitir que o usuário entre, saia e prossiga com seu fluxo de email.</span><span class="sxs-lookup"><span data-stu-id="0f96f-p106">Remember that the average Outlook session length on a phone is much shorter than on a PC. That means your add-in must be fast, and the scenario must allow the user to get in, get out, and get on with their email workflow.</span></span>

<span data-ttu-id="0f96f-130">Estes são exemplos de cenários que fazem sentido no Outlook Mobile.</span><span class="sxs-lookup"><span data-stu-id="0f96f-130">Here are examples of scenarios that make sense in Outlook Mobile.</span></span>

- <span data-ttu-id="0f96f-p107">O suplemento traz informações valiosas para o Outlook, para ajudar os usuários na triagem dos emails e a responder adequadamente. Exemplo: um suplemento CRM que permite ao usuário ver informações do cliente e compartilhar informações apropriadas.</span><span class="sxs-lookup"><span data-stu-id="0f96f-p107">The add-in brings valuable information into Outlook, helping users triage their email and respond appropriately. Example: a CRM add-in that lets the user see customer information and share appropriate information.</span></span>

- <span data-ttu-id="0f96f-p108">O suplemento agrega valor ao conteúdo do email do usuário, salvando as informações em um controle, uma colaboração ou um sistema semelhante. Exemplo: um suplemento que permite aos usuários ativar emails em itens de tarefa para acompanhamento de projetos, ou tíquetes de ajuda, para uma equipe de suporte.</span><span class="sxs-lookup"><span data-stu-id="0f96f-p108">The add-in adds value to the user's email content by saving the information to a tracking, collaboration, or similar system. Example: an add-in that lets users turn emails into task items for project tracking, or help tickets for a support team.</span></span>

<span data-ttu-id="0f96f-135">**Uma interação de usuário de exemplo para criar um cartão do Trello com base em uma mensagem de email no iOS**</span><span class="sxs-lookup"><span data-stu-id="0f96f-135">**An example user interaction to create a Trello card from an email message on iOS**</span></span>

![Um GIF animado mostrando a interação do usuário com um suplemento do Outlook Mobile no iOS](../images/outlook-mobile-addin-interaction.gif)

<br/>

<span data-ttu-id="0f96f-137">**Uma interação de usuário de exemplo para criar um cartão do Trello com base em uma mensagem de email no Android**</span><span class="sxs-lookup"><span data-stu-id="0f96f-137">**An example user interaction to create a Trello card from an email message on Android**</span></span>

![Um GIF animado mostrando a interação do usuário com um suplemento do Outlook Mobile no Android](../images/outlook-mobile-addin-interaction-android.gif)

## <a name="testing-your-add-ins-on-mobile"></a><span data-ttu-id="0f96f-139">Teste seus suplementos no celular</span><span class="sxs-lookup"><span data-stu-id="0f96f-139">Testing your add-ins on mobile</span></span>

<span data-ttu-id="0f96f-p109">Para testar um suplemento no Outlook Mobile, você pode carregar um suplemento para uma conta do O365 ou do Outlook.com. No Outlook na Web, acesse a engrenagem de configurações e escolha **Gerenciar Integrações** ou **Gerenciar Suplementos**. Perto da parte superior, clique onde diz: **Clique aqui para adicionar um suplemento personalizado** e carregue seu manifesto. Verifique se seu manifesto está formatado corretamente para conter `MobileFormFactor` ou ele não será carregado.</span><span class="sxs-lookup"><span data-stu-id="0f96f-p109">To test an add-in on Outlook Mobile, you can sideload an add-in to an O365 or Outlook.com account. In Outlook on the web, go to the settings gear, and choose **Manage Integrations** or **Manage Add-ins**. Near the top, click where it says **Click here to add a custom add-in** and upload your manifest. Make sure your manifest is properly formatted to contain `MobileFormFactor` or it won't load.</span></span>

<span data-ttu-id="0f96f-p110">Depois que seu suplemento estiver funcionando, certifique-se de testá-lo em tamanhos de tela diferentes, incluindo celulares e tablets. Você deve verificar se ele atende às diretrizes de acessibilidade de contraste, tamanho da fonte e cor, bem como de usabilidade com um leitor de tela, como o VoiceOver no iOS ou TalkBack no Android.</span><span class="sxs-lookup"><span data-stu-id="0f96f-p110">After your add-in is working, make sure to test it on different screen sizes, including phones and tablets. You should make sure it meets accessibility guidelines for contrast, font size, and color, as well as being usable with a screen reader such as VoiceOver on iOS or TalkBack on Android.</span></span>

<span data-ttu-id="0f96f-145">A solução de problemas no Mobile pode ser difícil, já que você pode não ter as ferramentas para as quais você está acostumado.</span><span class="sxs-lookup"><span data-stu-id="0f96f-145">Troubleshooting on mobile can be hard since you may not have the tools you're used to.</span></span> <span data-ttu-id="0f96f-146">No entanto, uma opção de solução de problemas no iOS é usar o Fiddler (Confira [este tutorial sobre como usá-lo com um dispositivo IOS](https://www.telerik.com/blogs/using-fiddler-with-apple-ios-devices)).</span><span class="sxs-lookup"><span data-stu-id="0f96f-146">However, one option for troubleshooting on iOS is to use Fiddler (check out [this tutorial on using it with an iOS device](https://www.telerik.com/blogs/using-fiddler-with-apple-ios-devices)).</span></span>

## <a name="next-steps"></a><span data-ttu-id="0f96f-147">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="0f96f-147">Next steps</span></span>

<span data-ttu-id="0f96f-148">Saiba como:</span><span class="sxs-lookup"><span data-stu-id="0f96f-148">Learn how to:</span></span>

- <span data-ttu-id="0f96f-149">[Adicionar suporte móvel ao manifesto do seu suplemento](add-mobile-support.md).</span><span class="sxs-lookup"><span data-stu-id="0f96f-149">[Add mobile support to your add-in's manifest](add-mobile-support.md).</span></span>
- <span data-ttu-id="0f96f-150">[Projetar uma ótima experiência móvel para seu suplemento](outlook-addin-design.md).</span><span class="sxs-lookup"><span data-stu-id="0f96f-150">[Design a great mobile experience for your add-in](outlook-addin-design.md).</span></span>
- <span data-ttu-id="0f96f-151">[Obter um token de acesso e chamar APIs REST do Outlook](use-rest-api.md) do suplemento.</span><span class="sxs-lookup"><span data-stu-id="0f96f-151">[Get an access token and call Outlook REST APIs](use-rest-api.md) from your add-in.</span></span>
