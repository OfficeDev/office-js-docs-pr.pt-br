---
title: Elemento IconUrl no arquivo de manifesto
description: ''
ms.date: 12/04/2018
ms.openlocfilehash: 471a168b5aa0091292132a1e078fa2b3f5efb448
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433121"
---
# <a name="iconurl-element"></a><span data-ttu-id="6f2fd-102">Elemento IconUrl</span><span class="sxs-lookup"><span data-stu-id="6f2fd-102">IconUrl element</span></span>

<span data-ttu-id="6f2fd-103">Especifica a URL da imagem que é usada para representar o seu Suplemento do Office na experiência de usuário de inserção e na Office Store.</span><span class="sxs-lookup"><span data-stu-id="6f2fd-103">Specifies the URL of the image that is used to represent your Office Add-in in the insertion UX and Office Store.</span></span>

<span data-ttu-id="6f2fd-104">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="6f2fd-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="6f2fd-105">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="6f2fd-105">Syntax</span></span>

```XML
<IconUrl DefaultValue="string" />
```

## <a name="can-contain"></a><span data-ttu-id="6f2fd-106">Pode conter</span><span class="sxs-lookup"><span data-stu-id="6f2fd-106">Can contain</span></span>

[<span data-ttu-id="6f2fd-107">Override</span><span class="sxs-lookup"><span data-stu-id="6f2fd-107">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="6f2fd-108">Atributos</span><span class="sxs-lookup"><span data-stu-id="6f2fd-108">Attributes</span></span>

|<span data-ttu-id="6f2fd-109">**Atributo**</span><span class="sxs-lookup"><span data-stu-id="6f2fd-109">**Attribute**</span></span>|<span data-ttu-id="6f2fd-110">**Tipo**</span><span class="sxs-lookup"><span data-stu-id="6f2fd-110">**Type**</span></span>|<span data-ttu-id="6f2fd-111">**Obrigatório**</span><span class="sxs-lookup"><span data-stu-id="6f2fd-111">**Required**</span></span>|<span data-ttu-id="6f2fd-112">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="6f2fd-112">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="6f2fd-113">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="6f2fd-113">DefaultValue</span></span>|<span data-ttu-id="6f2fd-114">cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="6f2fd-114">string</span></span>|<span data-ttu-id="6f2fd-115">obrigatório</span><span class="sxs-lookup"><span data-stu-id="6f2fd-115">required</span></span>|<span data-ttu-id="6f2fd-116">Especifica o valor padrão para essa configuração, expresso para a localidade especificada no elemento [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="6f2fd-116">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="6f2fd-117">Comentários</span><span class="sxs-lookup"><span data-stu-id="6f2fd-117">Remarks</span></span>

<span data-ttu-id="6f2fd-p101">Para um suplemento de email, o ícone é exibido na interface de usuário **Arquivo**  >  **Gerenciar suplementos** (Outlook) ou na interface de usuário **Configurações**  >  **Gerenciar suplementos** (Outlook Web App). Para um suplemento de conteúdo ou de painel de tarefas, o ícone é exibido na interface de usuário **Inserir**  >  **Suplementos**. Se você publicar o seu suplemento na Office Store, o ícone também será usado no site da Office Store para todos os tipos de suplementos.</span><span class="sxs-lookup"><span data-stu-id="6f2fd-p101">For a mail add-in, the icon is displayed in the  **File** > **Manage add-ins** UI (Outlook) or **Settings** > **Manage add-ins** UI (Outlook Web App). For a content or task pane add-in, the icon is displayed in the **Insert** > **Add-ins** UI. For all add-in types, the icon is also used on the Office Store site, if you publish your add-in to the Office Store.</span></span>

<span data-ttu-id="6f2fd-121">A imagem deve estar em um dos seguintes formatos: GIF, JPG, PNG, EXIF, BMP ou TIFF.</span><span class="sxs-lookup"><span data-stu-id="6f2fd-121">The image must be in one of the following file formats: GIF, JPG, PNG, EXIF, BMP, or TIFF.</span></span> <span data-ttu-id="6f2fd-122">Para aplicativos de conteúdo e de painel de tarefas, a imagem especificada deve ter 32 x 32 pixels.</span><span class="sxs-lookup"><span data-stu-id="6f2fd-122">For content and task pane apps, the image specified must be 32 x 32 pixels.</span></span> <span data-ttu-id="6f2fd-123">Para aplicativos de email, a resolução de imagem recomendada é 64 x 64 pixels.</span><span class="sxs-lookup"><span data-stu-id="6f2fd-123">For mail apps, the recommended image resolution is 64 x 64 pixels.</span></span> <span data-ttu-id="6f2fd-124">Você também deve especificar um ícone para ser usado com aplicativos host do Office executados em telas de DPI alto que utilizam o elemento [HighResolutionIconUrl](highresolutioniconurl.md).</span><span class="sxs-lookup"><span data-stu-id="6f2fd-124">You should also specify an icon for use with Office host applications running on high DPI screens using the [HighResolutionIconUrl](highresolutioniconurl.md) element.</span></span> <span data-ttu-id="6f2fd-125">Para saber mais, confira a seção _Criar uma identidade visual consistente para seu aplicativo_ em [Criar listagens eficazes no AppSource e no Office](https://docs.microsoft.com/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).</span><span class="sxs-lookup"><span data-stu-id="6f2fd-125">For more information, see the section _Create a consistent visual identity for your app_ in [Create effective listings in AppSource and within Office](https://docs.microsoft.com/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).</span></span>
