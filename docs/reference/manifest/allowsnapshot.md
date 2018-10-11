# <a name="allowsnapshot-element"></a><span data-ttu-id="6b23c-101">Elemento AllowSnapshot</span><span class="sxs-lookup"><span data-stu-id="6b23c-101">AllowSnapshot element</span></span>

<span data-ttu-id="6b23c-102">Especifica se uma imagem instantânea do seu suplemento de conteúdo é gravada com o documento host.</span><span class="sxs-lookup"><span data-stu-id="6b23c-102">Specifies whether a snapshot image of your content add-in is saved with the host document.</span></span>

<span data-ttu-id="6b23c-103">**Tipo de suplemento:** Conteúdo</span><span class="sxs-lookup"><span data-stu-id="6b23c-103">**Add-in type:** Content</span></span>

## <a name="syntax"></a><span data-ttu-id="6b23c-104">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="6b23c-104">Syntax</span></span>

```XML
<AllowSnapshot> [true | false]</AllowSnapshot>
```

## <a name="contained-in"></a><span data-ttu-id="6b23c-105">Contido em</span><span class="sxs-lookup"><span data-stu-id="6b23c-105">Contained in:</span></span>

[<span data-ttu-id="6b23c-106">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="6b23c-106">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="6b23c-107">Comentários</span><span class="sxs-lookup"><span data-stu-id="6b23c-107">Remarks</span></span>

 > [!IMPORTANT]
 > <span data-ttu-id="6b23c-108">**AllowSnapshot** é `true` por padrão.</span><span class="sxs-lookup"><span data-stu-id="6b23c-108">Security Note:**AllowSnapshot** is true`true` by default.</span></span> <span data-ttu-id="6b23c-109">Isso cria uma imagem do suplemento visível para os usuários que abrirem o documento em uma versão do aplicativo host que não oferece suporte a suplementos do Office, ou fornece uma imagem estática do suplemento se o aplicativo host não se conectar ao servidor que hospeda o suplemento.</span><span class="sxs-lookup"><span data-stu-id="6b23c-109">This makes an image of the add-in visible for users that open the document in a version of the host application that doesn't support Office Add-ins, or provides a static image of the add-in if the host application can't connect to the server hosting the add-in.</span></span> <span data-ttu-id="6b23c-110">No entanto, isso também significa que informações potencialmente confidenciais exibidas no suplemento podem ser acessadas diretamente a partir do documento que hospeda o suplemento.</span><span class="sxs-lookup"><span data-stu-id="6b23c-110">However, this also means that potentially sensitive information displayed in the add-in can be accessed directly from the document hosting the add-in.</span></span>

