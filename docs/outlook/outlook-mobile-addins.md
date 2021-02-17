---
title: Suplementos do Outlook para o Outlook Mobile
description: Os complementos do Outlook Mobile têm suporte em todas as contas comerciais do Microsoft 365, em Outlook.com e o suporte estará em breve nas contas do Gmail.
ms.date: 05/27/2020
localization_priority: Normal
ms.openlocfilehash: 586a473e1036e8480f395da49011f540d87e1b5f
ms.sourcegitcommit: 1cdf5728102424a46998e1527508b4e7f9f74a4c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/17/2021
ms.locfileid: "50270704"
---
# <a name="add-ins-for-outlook-mobile"></a><span data-ttu-id="9f05d-103">Suplementos do Outlook Mobile</span><span class="sxs-lookup"><span data-stu-id="9f05d-103">Add-ins for Outlook Mobile</span></span>

<span data-ttu-id="9f05d-p101">Os suplementos agora funcionam no Outlook Mobile, usando as mesmas APIs disponíveis para outros pontos de extremidade do Outlook. Se você já tiver criado um suplemento para Outlook, é fácil fazê-lo funcionar no Outlook Mobile.</span><span class="sxs-lookup"><span data-stu-id="9f05d-p101">Add-ins now work on Outlook Mobile, using the same APIs available for other Outlook endpoints. If you've built an add-in for Outlook already, it's easy to get it working on Outlook Mobile.</span></span>

<span data-ttu-id="9f05d-106">Os complementos do Outlook Mobile têm suporte em todas as contas comerciais do Microsoft 365, Outlook.com e o suporte estará chegando em breve às contas do Gmail.</span><span class="sxs-lookup"><span data-stu-id="9f05d-106">Outlook mobile add-ins are supported on all Microsoft 365 business accounts, Outlook.com accounts, and support is coming soon to Gmail accounts.</span></span>

<span data-ttu-id="9f05d-107">**Um painel de tarefas de exemplo no Outlook no iOS**</span><span class="sxs-lookup"><span data-stu-id="9f05d-107">**An example task pane in Outlook on iOS**</span></span>

![Uma captura de tela do painel de tarefas no Outlook no iOS](../images/outlook-mobile-addin-taskpane.png)

<br/>

<span data-ttu-id="9f05d-109">**Um painel de tarefas de exemplo no Outlook no Android**</span><span class="sxs-lookup"><span data-stu-id="9f05d-109">**An example task pane in Outlook on Android**</span></span>

![Uma captura de tela do painel de tarefas no Outlook no Android](../images/outlook-mobile-addin-taskpane-android.png)

> [!IMPORTANT]
> <span data-ttu-id="9f05d-111">Os complementos não funcionam na versão moderna do Outlook em um navegador móvel.</span><span class="sxs-lookup"><span data-stu-id="9f05d-111">Add-ins don't work in the modern version of Outlook in a mobile browser.</span></span> <span data-ttu-id="9f05d-112">Para saber mais, confira [o Outlook em seu navegador móvel que está sendo atualizado.](https://techcommunity.microsoft.com/t5/outlook-blog/outlook-on-your-mobile-browser-is-being-upgraded/ba-p/1125816)</span><span class="sxs-lookup"><span data-stu-id="9f05d-112">For more information, see [Outlook on your mobile browser is being upgraded](https://techcommunity.microsoft.com/t5/outlook-blog/outlook-on-your-mobile-browser-is-being-upgraded/ba-p/1125816).</span></span>

## <a name="whats-different-on-mobile"></a><span data-ttu-id="9f05d-113">Qual é a diferença no celular?</span><span class="sxs-lookup"><span data-stu-id="9f05d-113">What's different on mobile?</span></span>

- <span data-ttu-id="9f05d-p103">O tamanho pequeno e as rápidas interações tornam o projeto para celular um desafio. Para garantir experiências de qualidade para nossos clientes, estamos definindo critérios rígidos de validação que devem ser cumpridos por um suplemento que declara suporte a celular de forma a ser aprovado na AppSource.</span><span class="sxs-lookup"><span data-stu-id="9f05d-p103">The small size and quick interactions make designing for mobile a challenge. To ensure quality experiences for our customers, we are setting strict validation criteria that must be met by an add-in declaring mobile support, in order to be approved in AppSource.</span></span>
  - <span data-ttu-id="9f05d-116">O suplemento **DEVE** cumprir as [diretrizes de interface do usuário](outlook-addin-design.md).</span><span class="sxs-lookup"><span data-stu-id="9f05d-116">The add-in **MUST** adhere to the [UI guidelines](outlook-addin-design.md).</span></span>
  - <span data-ttu-id="9f05d-117">O cenário do suplemento **DEVE** [fazer sentido no mobile](#what-makes-a-good-scenario-for-mobile-add-ins).</span><span class="sxs-lookup"><span data-stu-id="9f05d-117">The scenario for the add-in **MUST** [make sense on mobile](#what-makes-a-good-scenario-for-mobile-add-ins).</span></span>

- <span data-ttu-id="9f05d-118">Em geral, somente o modo De leitura de mensagem é suportado no momento.</span><span class="sxs-lookup"><span data-stu-id="9f05d-118">In general, only Message Read mode is supported at this time.</span></span> <span data-ttu-id="9f05d-119">Isso significa `MobileMessageReadCommandSurface` que é o único [ExtensionPoint](../reference/manifest/extensionpoint.md#mobilemessagereadcommandsurface) que você deve declarar na seção móvel do manifesto.</span><span class="sxs-lookup"><span data-stu-id="9f05d-119">That means `MobileMessageReadCommandSurface` is the only [ExtensionPoint](../reference/manifest/extensionpoint.md#mobilemessagereadcommandsurface) you should declare in the mobile section of your manifest.</span></span> <span data-ttu-id="9f05d-120">No entanto, o modo Organizador de Compromissos é suportado para os complementos integrados do provedor de reunião online que, em vez disso, declaram o ponto de extensão [MobileOnlineMeetingCommandSurface](../reference/manifest/extensionpoint.md#mobileonlinemeetingcommandsurface).</span><span class="sxs-lookup"><span data-stu-id="9f05d-120">However, Appointment Organizer mode is supported for online meeting provider integrated add-ins which instead declare the [MobileOnlineMeetingCommandSurface extension point](../reference/manifest/extensionpoint.md#mobileonlinemeetingcommandsurface).</span></span> <span data-ttu-id="9f05d-121">Consulte o [artigo Criar um complemento móvel do Outlook para um provedor](online-meeting.md) de reuniões online para saber mais sobre esse cenário.</span><span class="sxs-lookup"><span data-stu-id="9f05d-121">See the [Create an Outlook mobile add-in for an online-meeting provider](online-meeting.md) article for more about this scenario.</span></span>

- <span data-ttu-id="9f05d-p105">A API [makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) não é suportada no celular, já que o aplicativo móvel usa APIs REST para se comunicar com o servidor. Se seu back-end do aplicativo precisa se conectar ao servidor do Exchange, é possível usar o token de retorno de chamada para fazer chamadas de API REST. Para obter detalhes, consulte [Usar APIs REST do Outlook de um suplemento do Outlook](use-rest-api.md).</span><span class="sxs-lookup"><span data-stu-id="9f05d-p105">The [makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) API is not supported on mobile since the mobile app uses REST APIs to communicate with the server. If your app backend needs to connect to the Exchange server, you can use the callback token to make REST API calls. For details, see [Use the Outlook REST APIs from an Outlook add-in](use-rest-api.md).</span></span>

- <span data-ttu-id="9f05d-125">Quando você envia o suplemento para a loja com [MobileFormFactor](../reference/manifest/mobileformfactor.md) no manifesto, precisará concordar com nosso adendo de suplementos no iOS e precisará enviar sua ID de desenvolvedor Apple para verificação.</span><span class="sxs-lookup"><span data-stu-id="9f05d-125">When you submit your add-in to the store with [MobileFormFactor](../reference/manifest/mobileformfactor.md) in the manifest, you'll need to agree to our developer addendum for add-ins on iOS, and you must submit your Apple Developer ID for verification.</span></span>

- <span data-ttu-id="9f05d-126">Por fim, seu manifesto precisará declarar `MobileFormFactor` e ter os tipos corretos de [controles](../reference/manifest/control.md) e [tamanhos de ícone](../reference/manifest/icon.md) incluídos.</span><span class="sxs-lookup"><span data-stu-id="9f05d-126">Finally, your manifest will need to declare `MobileFormFactor`, and have the correct types of [controls](../reference/manifest/control.md) and [icon sizes](../reference/manifest/icon.md) included.</span></span>

## <a name="what-makes-a-good-scenario-for-mobile-add-ins"></a><span data-ttu-id="9f05d-127">O que forma um bom cenário para suplementos móveis?</span><span class="sxs-lookup"><span data-stu-id="9f05d-127">What makes a good scenario for mobile add-ins?</span></span>

<span data-ttu-id="9f05d-p106">Lembre-se de que o tamanho médio da sessão Outlook em um telefone é bem menor do que em um PC. Isso significa que seu suplemento deve ser rápido e o cenário deve permitir que o usuário entre, saia e prossiga com seu fluxo de email.</span><span class="sxs-lookup"><span data-stu-id="9f05d-p106">Remember that the average Outlook session length on a phone is much shorter than on a PC. That means your add-in must be fast, and the scenario must allow the user to get in, get out, and get on with their email workflow.</span></span>

<span data-ttu-id="9f05d-130">Estes são exemplos de cenários que fazem sentido no Outlook Mobile.</span><span class="sxs-lookup"><span data-stu-id="9f05d-130">Here are examples of scenarios that make sense in Outlook Mobile.</span></span>

- <span data-ttu-id="9f05d-p107">O suplemento traz informações valiosas para o Outlook, para ajudar os usuários na triagem dos emails e a responder adequadamente. Exemplo: um suplemento CRM que permite ao usuário ver informações do cliente e compartilhar informações apropriadas.</span><span class="sxs-lookup"><span data-stu-id="9f05d-p107">The add-in brings valuable information into Outlook, helping users triage their email and respond appropriately. Example: a CRM add-in that lets the user see customer information and share appropriate information.</span></span>

- <span data-ttu-id="9f05d-p108">O suplemento agrega valor ao conteúdo do email do usuário, salvando as informações em um controle, uma colaboração ou um sistema semelhante. Exemplo: um suplemento que permite aos usuários ativar emails em itens de tarefa para acompanhamento de projetos, ou tíquetes de ajuda, para uma equipe de suporte.</span><span class="sxs-lookup"><span data-stu-id="9f05d-p108">The add-in adds value to the user's email content by saving the information to a tracking, collaboration, or similar system. Example: an add-in that lets users turn emails into task items for project tracking, or help tickets for a support team.</span></span>

<span data-ttu-id="9f05d-135">**Uma interação de usuário de exemplo para criar um cartão do Trello com base em uma mensagem de email no iOS**</span><span class="sxs-lookup"><span data-stu-id="9f05d-135">**An example user interaction to create a Trello card from an email message on iOS**</span></span>

![Um GIF animado mostrando a interação do usuário com um suplemento do Outlook Mobile no iOS](../images/outlook-mobile-addin-interaction.gif)

<br/>

<span data-ttu-id="9f05d-137">**Uma interação de usuário de exemplo para criar um cartão do Trello com base em uma mensagem de email no Android**</span><span class="sxs-lookup"><span data-stu-id="9f05d-137">**An example user interaction to create a Trello card from an email message on Android**</span></span>

![Um GIF animado mostrando a interação do usuário com um suplemento do Outlook Mobile no Android](../images/outlook-mobile-addin-interaction-android.gif)

## <a name="testing-your-add-ins-on-mobile"></a><span data-ttu-id="9f05d-139">Teste seus suplementos no celular</span><span class="sxs-lookup"><span data-stu-id="9f05d-139">Testing your add-ins on mobile</span></span>

<span data-ttu-id="9f05d-p109">Para testar um suplemento no Outlook Mobile, você pode carregar um suplemento para uma conta do O365 ou do Outlook.com. No Outlook na Web, acesse a engrenagem de configurações e escolha **Gerenciar Integrações** ou **Gerenciar Suplementos**. Perto da parte superior, clique onde diz: **Clique aqui para adicionar um suplemento personalizado** e carregue seu manifesto. Verifique se seu manifesto está formatado corretamente para conter `MobileFormFactor` ou ele não será carregado.</span><span class="sxs-lookup"><span data-stu-id="9f05d-p109">To test an add-in on Outlook Mobile, you can sideload an add-in to an O365 or Outlook.com account. In Outlook on the web, go to the settings gear, and choose **Manage Integrations** or **Manage Add-ins**. Near the top, click where it says **Click here to add a custom add-in** and upload your manifest. Make sure your manifest is properly formatted to contain `MobileFormFactor` or it won't load.</span></span>

<span data-ttu-id="9f05d-p110">Depois que seu suplemento estiver funcionando, certifique-se de testá-lo em tamanhos de tela diferentes, incluindo celulares e tablets. Você deve verificar se ele atende às diretrizes de acessibilidade de contraste, tamanho da fonte e cor, bem como de usabilidade com um leitor de tela, como o VoiceOver no iOS ou TalkBack no Android.</span><span class="sxs-lookup"><span data-stu-id="9f05d-p110">After your add-in is working, make sure to test it on different screen sizes, including phones and tablets. You should make sure it meets accessibility guidelines for contrast, font size, and color, as well as being usable with a screen reader such as VoiceOver on iOS or TalkBack on Android.</span></span>

<span data-ttu-id="9f05d-145">A solução de problemas em dispositivos móveis pode ser difícil, pois talvez você não tenha as ferramentas com as que está acostumado.</span><span class="sxs-lookup"><span data-stu-id="9f05d-145">Troubleshooting on mobile can be hard since you may not have the tools you're used to.</span></span> <span data-ttu-id="9f05d-146">No entanto, uma opção para solucionar problemas no iOS é usar o Fiddler (confira este tutorial sobre como [usá-lo com um dispositivo iOS).](https://www.telerik.com/blogs/using-fiddler-with-apple-ios-devices)</span><span class="sxs-lookup"><span data-stu-id="9f05d-146">However, one option for troubleshooting on iOS is to use Fiddler (check out [this tutorial on using it with an iOS device](https://www.telerik.com/blogs/using-fiddler-with-apple-ios-devices)).</span></span>

## <a name="next-steps"></a><span data-ttu-id="9f05d-147">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="9f05d-147">Next steps</span></span>

<span data-ttu-id="9f05d-148">Saiba como:</span><span class="sxs-lookup"><span data-stu-id="9f05d-148">Learn how to:</span></span>

- <span data-ttu-id="9f05d-149">[Adicionar suporte móvel ao manifesto do seu suplemento](add-mobile-support.md).</span><span class="sxs-lookup"><span data-stu-id="9f05d-149">[Add mobile support to your add-in's manifest](add-mobile-support.md).</span></span>
- <span data-ttu-id="9f05d-150">[Projetar uma ótima experiência móvel para seu suplemento](outlook-addin-design.md).</span><span class="sxs-lookup"><span data-stu-id="9f05d-150">[Design a great mobile experience for your add-in](outlook-addin-design.md).</span></span>
- <span data-ttu-id="9f05d-151">[Obter um token de acesso e chamar APIs REST do Outlook](use-rest-api.md) do suplemento.</span><span class="sxs-lookup"><span data-stu-id="9f05d-151">[Get an access token and call Outlook REST APIs](use-rest-api.md) from your add-in.</span></span>
