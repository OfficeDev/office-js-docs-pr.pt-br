---
title: Ciclo de vida de desenvolvimento de suplementos do Office
description: ''
ms.date: 07/01/2019
localization_priority: Priority
ms.openlocfilehash: 44e2792f030662bd89b272998ad47fd0a645d785
ms.sourcegitcommit: 90c2d8236c6b30d80ac2b13950028a208ef60973
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/02/2019
ms.locfileid: "35454570"
---
# <a name="office-add-ins-development-lifecycle"></a><span data-ttu-id="30e6e-102">Ciclo de vida de desenvolvimento de suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="30e6e-102">Office Add-ins development lifecycle</span></span>

> [!NOTE]
> <span data-ttu-id="30e6e-p101">Caso pretenda [publicar](../publish/publish.md) o suplemento na experiência do Office depois de criá-lo, verifique se você está em conformidade com as [Políticas de validação do AppSource](/office/dev/store/validation-policies). Por exemplo, para passar na validação, seu suplemento deve funcionar em todas as plataformas com suporte aos métodos que você definir (para mais informações, confira a [seção 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) e a [Página de hospedagem e disponibilidade de suplementos do Office](../overview/office-add-in-availability.md)).</span><span class="sxs-lookup"><span data-stu-id="30e6e-p101">If you plan to [publish](../publish/publish.md) your add-in to AppSource and make it available within the Office experience, make sure that you conform to the [AppSource validation policies](/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)).</span></span> 

<span data-ttu-id="30e6e-105">O ciclo de vida de desenvolvimento típico de um Suplemento do Office inclui as seguintes etapas:</span><span class="sxs-lookup"><span data-stu-id="30e6e-105">The typical development lifecycle of an Office Add-in includes the following steps:</span></span>


## <a name="1-decide-on-the-purpose-of-the-add-in"></a><span data-ttu-id="30e6e-106">1. Decida qual é a proposta do suplemento</span><span class="sxs-lookup"><span data-stu-id="30e6e-106">1. Decide on the purpose of the add-in</span></span>

<span data-ttu-id="30e6e-107">Faça as seguintes perguntas:</span><span class="sxs-lookup"><span data-stu-id="30e6e-107">Ask the following questions:</span></span>

- <span data-ttu-id="30e6e-108">Para quê o suplemento é útil?</span><span class="sxs-lookup"><span data-stu-id="30e6e-108">How is the add-in useful?</span></span>

- <span data-ttu-id="30e6e-109">Como ele ajuda seus clientes a serem mais produtivos?</span><span class="sxs-lookup"><span data-stu-id="30e6e-109">How does it help your customers be more productive?</span></span>

- <span data-ttu-id="30e6e-110">Quais cenários são compatíveis com os recursos do seu suplemento?</span><span class="sxs-lookup"><span data-stu-id="30e6e-110">What scenarios does your add-in's features support?</span></span>

<span data-ttu-id="30e6e-111">Decida os recursos e cenários mais importantes e concentre seu design nisso.</span><span class="sxs-lookup"><span data-stu-id="30e6e-111">Decide the most important features and scenarios and focus your design around them.</span></span>


## <a name="2-identify-the-data-and-data-source-for-the-add-in"></a><span data-ttu-id="30e6e-112">2. Identifique os dados e a fonte de dados do suplemento</span><span class="sxs-lookup"><span data-stu-id="30e6e-112">2. Identify the data and data source for the add-in</span></span>

- <span data-ttu-id="30e6e-113">Os dados estão em um documento, uma pasta de trabalho, uma apresentação ou um projeto?</span><span class="sxs-lookup"><span data-stu-id="30e6e-113">Is the data in a document, workbook, presentation, project, or an Access browser-based database?</span></span>

- <span data-ttu-id="30e6e-114">Os dados sobre um item ou itens estão no Exchange Server ou em uma caixa de correio do Exchange Online?</span><span class="sxs-lookup"><span data-stu-id="30e6e-114">Is the data about an item or items in an Exchange Server or Exchange Online mailbox?</span></span>

- <span data-ttu-id="30e6e-115">Os dados são provenientes de uma fonte externa, como um serviço Web?</span><span class="sxs-lookup"><span data-stu-id="30e6e-115">Is the data from an external source such as a web service?</span></span>


## <a name="3-identify-the-type-of-add-in-and-office-host-applications-that-best-support-the-purpose-of-the-add-in"></a><span data-ttu-id="30e6e-116">3. Identifique o tipo de suplemento e os aplicativos host do Office que dão o melhor suporte à finalidade do suplemento.</span><span class="sxs-lookup"><span data-stu-id="30e6e-116">3. Identify the type of add-in and Office host applications that best support the purpose of the add-in</span></span>

<span data-ttu-id="30e6e-117">Considere o seguinte para identificar os cenários:</span><span class="sxs-lookup"><span data-stu-id="30e6e-117">Consider the following to identify the scenarios:</span></span>

- <span data-ttu-id="30e6e-p102">Os clientes usarão o suplemento para enriquecer o conteúdo de um documento? Em caso afirmativo, convém considerar a criação de um **suplemento de conteúdo**.</span><span class="sxs-lookup"><span data-stu-id="30e6e-p102">Will customers use the add-in to enrich the content of a document or Access browser-based database? If so, you may want to consider creating a **content add-in**.</span></span>

- <span data-ttu-id="30e6e-p103">Os clientes utilizarão o suplemento ao exibir ou ao escrever uma mensagem de email ou um compromisso? É importante poder expor o suplemento de acordo com o contexto atual? É uma prioridade disponibilizar o suplemento não apenas em computadores de mesa, mas também em tablets e telefones?</span><span class="sxs-lookup"><span data-stu-id="30e6e-p103">Will customers use the add-in while viewing or composing an email message or appointment? Is being able to expose the add-in according to the current context important? Is making the add-in available on not just the desktop, but also on tablets and phones a priority?</span></span>

    <span data-ttu-id="30e6e-p104">Se a resposta for Sim para qualquer uma dessas perguntas, considere a criação de um **suplemento do Outlook**. Identifique o contexto que acionará seu suplemento (por exemplo, o usuário está usando um formulário de composição, tipos de mensagem específicos, a presença de um anexo, um endereço, uma sugestão de tarefa ou de reunião, ou certos padrões de cadeia de caracteres no conteúdo de um compromisso ou um email).</span><span class="sxs-lookup"><span data-stu-id="30e6e-p104">If you answer yes to any of these questions, consider creating an **Outlook add-in**. Identify the context that will trigger your add-in (for example, the user being in a compose form, specific message types, the presence of an attachment, address, task suggestion, or meeting suggestion, or certain string patterns in the contents of an email or appointment).</span></span>

    <span data-ttu-id="30e6e-125">Para descobrir como é possível ativar o suplemento Outlook contextualmente, confira as [Regras de ativação para suplementos do Outlook](/outlook/add-ins/activation-rules).</span><span class="sxs-lookup"><span data-stu-id="30e6e-125">To find out how you can contextually activate the Outlook add-in, see [Activation rules for Outlook add-ins](/outlook/add-ins/activation-rules).</span></span>

- <span data-ttu-id="30e6e-p105">Os clientes usarão o suplemento para aprimorar a experiência de criação ou de exibição de um documento? Em caso afirmativo, convém considerar a criação de um **suplemento de painel de tarefas**.</span><span class="sxs-lookup"><span data-stu-id="30e6e-p105">Will customers use the add-in to enhance the viewing or authoring experience of a document? If so, you may want to consider creating a **task pane add-in**.</span></span>

<span data-ttu-id="30e6e-128">O suporte para determinadas APIs de suplemento pode ser diferente entre aplicativos do Office e de acordo com a plataforma em que estão sendo executados (no Windows, em Mac, na Web ou em dispositivos móveis).</span><span class="sxs-lookup"><span data-stu-id="30e6e-128">Support for certain add-in APIs may differ between Office applications and the platform they are running on (Windows, Mac, Web, Mobile).</span></span> <span data-ttu-id="30e6e-129">Para ver a cobertura da API atual pelo cliente e a plataforma, consulte nossa página [Disponibilidade de plataforma e host para o Suplemento do Office](../overview/office-add-in-availability.md).</span><span class="sxs-lookup"><span data-stu-id="30e6e-129">To see the current API coverage by client and platform, see our [Office Add-in host and platform availability](../overview/office-add-in-availability.md) page.</span></span>  


## <a name="4-design-and-implement-the-user-experience-and-user-interface-for-the-add-in"></a><span data-ttu-id="30e6e-130">4. Desenvolva e implemente a experiência do usuário e a interface do usuário para o suplemento.</span><span class="sxs-lookup"><span data-stu-id="30e6e-130">4. Design and implement the user experience and user interface for the add-in</span></span>

<span data-ttu-id="30e6e-p107">Projete uma experiência de usuário rápida e fluida, que seja consistente, fácil de usar e com cenários primários que requerem apenas algumas etapas para serem executados. Dependendo da finalidade do suplemento, use APIs ou serviços da Web de terceiros.</span><span class="sxs-lookup"><span data-stu-id="30e6e-p107">Design a fast and fluid user experience that is consistent, easy to learn, with primary scenarios that require only a few steps to complete. Depending on the purpose of the add-in, make use of third-party APIs or web services.</span></span>

<span data-ttu-id="30e6e-133">Você pode escolher entre várias ferramentas de desenvolvimento na Web e usar o HTML e JavaScript para implementar a interface do usuário.</span><span class="sxs-lookup"><span data-stu-id="30e6e-133">You can choose from a variety of web development tools and use HTML and JavaScript to implement the user interface.</span></span>


## <a name="5-create-an-xml-manifest-file-based-on-the-office-add-ins-manifest-schema"></a><span data-ttu-id="30e6e-134">5. Crie um arquivo de manifesto XML com base no esquema do manifesto dos Suplementos do Office.</span><span class="sxs-lookup"><span data-stu-id="30e6e-134">5. Create an XML manifest file based on the Office Add-ins manifest schema</span></span>

<span data-ttu-id="30e6e-135">Crie um manifesto XML para identificar o suplemento e seus requisitos, especificar os locais do HTML e de arquivos JavaScript e CSS que o suplemento possa vir a usar e, dependendo do tipo de suplemento, o tamanho e as permissões padrão.</span><span class="sxs-lookup"><span data-stu-id="30e6e-135">Create an XML manifest to identify the add-in and its requirements, specify the locations of the HTML and any JavaScript and CSS files that the add-in uses, and depending on the type of the add-in, the default size and permissions.</span></span>

<span data-ttu-id="30e6e-p108">Para suplementos do Outlook, é possível especificar o contexto (com base na mensagem ou no compromisso atual) relevante para seu suplemento e que, portanto, faria o Outlook disponibilizá-lo na interface de usuário. Também é possível decidir quais dispositivos serão compatíveis com o suplemento. No manifesto, especifique o contexto para regras de ativação e dispositivos compatíveis.</span><span class="sxs-lookup"><span data-stu-id="30e6e-p108">For Outlook add-ins, you can specify the context, based on the current message or appointment, under which your add-in is relevant and you would like Outlook to make available in the UI. You can also decide which devices you want the add-in to support. In the manifest, specify the context as activation rules and the supported devices.</span></span>


## <a name="6-install-and-test-the-add-in"></a><span data-ttu-id="30e6e-139">6. Instale e teste o suplemento.</span><span class="sxs-lookup"><span data-stu-id="30e6e-139">6. Install and test the add-in</span></span>

<span data-ttu-id="30e6e-p109">Coloque os arquivos HTML e todos os arquivos JavaScript e CSS nos servidores Web especificados no arquivo de manifesto do suplemento. O processo de instalação de um suplemento depende do tipo de suplemento. Para obter detalhes, confira [Realizar Sideload de Suplementos do Office para teste](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="30e6e-p109">Place the HTML files and any JavaScript and CSS files on the web servers that are specified in the add-in manifest file. The process to install an add-in depends on the type of the add-in. For details, see [Sideload Office Add-ins for testing](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span></span>

<span data-ttu-id="30e6e-p110">Para suplementos do Outlook, instale-os em uma caixa de correio do Exchange e especifique o local do arquivo de manifesto do suplemento no Centro de Administração do Exchange (EAC). Para saber mais, consulte [Implementar e instalar suplementos do Outlook para teste](/outlook/add-ins/testing-and-tips).</span><span class="sxs-lookup"><span data-stu-id="30e6e-p110">For Outlook add-ins, install it in an Exchange mailbox, and specify the location of the add-in manifest file in the Exchange Admin Center (EAC). For more information, see [Deploy and install Outlook add-ins for testing](/outlook/add-ins/testing-and-tips).</span></span>


## <a name="7-publish-the-add-in"></a><span data-ttu-id="30e6e-145">7. Publique o suplemento.</span><span class="sxs-lookup"><span data-stu-id="30e6e-145">7. Publish the add-in</span></span>

<span data-ttu-id="30e6e-p111">Você pode enviar o suplemento para o AppSource, a partir do qual os clientes podem instalar o suplemento. Além disso, você pode publicar os suplementos de painel de tarefas e de conteúdo em um catálogo de aplicativos no SharePoint ou em uma pasta de rede compartilhada e pode implantar um suplemento do Outlook diretamente em um Exchange Server para a sua organização. Para detalhes, consulte [Publicar seu Suplemento do Office](../publish/publish.md).</span><span class="sxs-lookup"><span data-stu-id="30e6e-p111">You can submit the add-in to AppSource, from which customers can install the add-in. In addition, you can publish task pane and content add-ins to a private folder add-in catalog on SharePoint or to a shared network folder, and you can deploy an Outlook add-in directly on an Exchange server for your organization. For details, see [Publish your Office Add-in](../publish/publish.md).</span></span>


## <a name="8-maintain-the-add-in"></a><span data-ttu-id="30e6e-149">8. Faça a manutenção do suplemento</span><span class="sxs-lookup"><span data-stu-id="30e6e-149">8. Maintain the add-in</span></span>

<span data-ttu-id="30e6e-p112">Se o suplemento chamar um serviço Web e você atualizar o serviço Web depois de publicar o suplemento, não será preciso publicar o suplemento novamente. No entanto, se você alterar os itens ou dados enviados ao suplemento, por exemplo, o manifesto do suplemento, capturas de tela, ícones, arquivos HTML ou JavaScript, você precisará publicá-lo novamente.</span><span class="sxs-lookup"><span data-stu-id="30e6e-p112">If your add-in calls a web service, and if you make updates to the web service after publishing the add-in, you do not have to republish the add-in. However, if you change any items or data you submitted for your add-in, such as the add-in manifest, screenshots, icons, HTML or JavaScript files, you will need to republish the add-in.</span></span> 

<span data-ttu-id="30e6e-p113">Especificamente, se você publicar o suplemento no AppSource, será preciso reenviar o suplemento para que o AppSource possa implementar as alterações. Você deve reenviar o suplemento com o manifesto de suplemento atualizado que inclui um novo número da versão. Você também deve se certificar de atualizar o número da versão do suplemento no formulário de envio para corresponder ao novo número da versão do manifesto. Para suplementos do Outlook, verifique se o elemento [Id](/office/dev/add-ins/reference/manifest/id) contém um UUID diferente do manifesto de suplemento.</span><span class="sxs-lookup"><span data-stu-id="30e6e-p113">In particular, if you have published the add-in to AppSource, you'll need to resubmit your add-in so that AppSource can implement those changes. You must resubmit your add-in with an updated add-in manifest that includes a new version number. You must also make sure to update the add-in version number in the submission form to match the new manifest's version number. For Outlook add-ins, you should make sure the [Id](/office/dev/add-ins/reference/manifest/id) element contains a different UUID in the add-in manifest.</span></span>
