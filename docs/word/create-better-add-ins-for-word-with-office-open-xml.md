---
title: Criar suplementos melhores para o Word com o Office Open XML
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: e13911da0dbdb9fdb0215d433a9559bf1b747eb9
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32449986"
---
# <a name="create-better-add-ins-for-word-with-office-open-xml"></a><span data-ttu-id="79fca-102">Criar suplementos melhores para o Word com o Office Open XML</span><span class="sxs-lookup"><span data-stu-id="79fca-102">Create better add-ins for Word with Office Open XML</span></span>

<span data-ttu-id="79fca-103">**Fornecido por:** Stephanie Krieger, Microsoft Corporation | Juan Balmori Labra, Microsoft Corporation</span><span class="sxs-lookup"><span data-stu-id="79fca-103">**Provided by:** Stephanie Krieger, Microsoft Corporation | Juan Balmori Labra, Microsoft Corporation</span></span>

<span data-ttu-id="79fca-p101">Se você está criando suplementos do Office para serem executados no Word, talvez já saiba que a API JavaScript para Office (Office.js) oferece vários formatos para ler e gravar o conteúdo de documentos. Eles são chamados de tipos de coerção e incluem texto sem formatação, tabelas, HTML e Office Open XML.</span><span class="sxs-lookup"><span data-stu-id="79fca-p101">If you're building Office Add-ins to run in Word, you might already know that the JavaScript API for Office (Office.js) offers several formats for reading and writing document content. These are called coercion types, and they include plain text, tables, HTML, and Office Open XML.</span></span>

<span data-ttu-id="79fca-p102">Então, quais são suas opções quando você precisa adicionar conteúdo avançado a um documento, como imagens, tabelas formatadas, gráficos ou apenas texto formatado? Você pode usar HTML para inserir alguns tipos de conteúdo avançado, como imagens. Dependendo do cenário, pode haver desvantagens na coerção de HTML, como limitações nas opções de formatação e posicionamento disponíveis para o conteúdo. Como o Office Open XML é a linguagem na qual os documentos do Word (como .docx e .dotx) são gravados, você pode inserir praticamente qualquer tipo de conteúdo que um usuário pode adicionar a um documento do Word, com praticamente qualquer tipo de formatação que o usuário possa aplicar. Determinar a marcação do Office Open XML necessária para fazer isso é mais fácil do que você imagina.</span><span class="sxs-lookup"><span data-stu-id="79fca-p102">So what are your options when you need to add rich content to a document, such as images, formatted tables, charts, or even just formatted text? You can use HTML for inserting some types of rich content, such as pictures. Depending on your scenario, there can be drawbacks to HTML coercion, such as limitations in the formatting and positioning options available to your content. Because Office Open XML is the language in which Word documents (such as .docx and .dotx) are written, you can insert virtually any type of content that a user can add to a Word document, with virtually any type of formatting the user can apply. Determining the Office Open XML markup you need to get it done is easier than you might think.</span></span>

> [!NOTE]
> <span data-ttu-id="79fca-p103">O Office Open XML também é a linguagem por trás dos documentos do PowerPoint e do Excel (e, a partir do Office 2013, do Visio). No entanto, atualmente, você pode fazer a coerção de conteúdo como Office Open XML somente em Suplementos do Office criados para o Word. Para saber mais sobre o Office Open XML, incluindo a documentação de referência completa da linguagem, confira [Recursos adicionais](#see-also).</span><span class="sxs-lookup"><span data-stu-id="79fca-p103">Office Open XML is also the language behind PowerPoint and Excel (and, as of Office 2013, Visio) documents. However, currently, you can coerce content as Office Open XML only in Office Add-ins created for Word. For more information about Office Open XML, including the complete language reference documentation, see [Additional resources](#see-also).</span></span>

<span data-ttu-id="79fca-p104">Para começar, veja alguns dos tipos de conteúdo que você pode inserir usando a coerção do Office Open XML. Baixe o exemplo de código [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML), que contém a marcação do Office Open XML e o código Office.js necessário para inserir qualquer um dos exemplos a seguir no Word.</span><span class="sxs-lookup"><span data-stu-id="79fca-p104">To begin, take a look at some of the content types you can insert using Office Open XML coercion. Download the code sample [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML), which contains the Office Open XML markup and Office.js code required for inserting any of the following examples into Word.</span></span>

> [!NOTE]
> <span data-ttu-id="79fca-116">Ao longo deste artigo, os termos **tipos de conteúdo** e **conteúdo avançado** referem-se aos tipos de conteúdo avançado que você pode inserir em um documento do Word.</span><span class="sxs-lookup"><span data-stu-id="79fca-116">Throughout this article, the terms  **content types** and **rich content** refer to the types of rich content you can insert into a Word document.</span></span>


<span data-ttu-id="79fca-117">*Figura 1. Texto com formatação direta*</span><span class="sxs-lookup"><span data-stu-id="79fca-117">*Figure 1. Text with direct formatting*</span></span>


![Texto com formatação direta aplicada.](../images/office15-app-create-wd-app-using-ooxml-fig01.png)

<span data-ttu-id="79fca-119">Você pode usar a formatação direta para especificar a aparência exata que o texto terá, independentemente da formatação existente no documento do usuário.</span><span class="sxs-lookup"><span data-stu-id="79fca-119">You can use direct formatting to specify exactly what the text will look like regardless of existing formatting in the user's document.</span></span>

<span data-ttu-id="79fca-120">*Figura 2. Texto formatado com um estilo*</span><span class="sxs-lookup"><span data-stu-id="79fca-120">*Figure 2. Text formatted using a style*</span></span>


![Texto formatado com estilo de parágrafo.](../images/office15-app-create-wd-app-using-ooxml-fig02.png)

<span data-ttu-id="79fca-122">Você pode usar um estilo para coordenar automaticamente a aparência do texto que insere com o documento do usuário.</span><span class="sxs-lookup"><span data-stu-id="79fca-122">You can use a style to automatically coordinate the look of text you insert with the user's document.</span></span>

<span data-ttu-id="79fca-123">*Figura 3. Uma imagem simples*</span><span class="sxs-lookup"><span data-stu-id="79fca-123">*Figure 3. A simple image*</span></span>


![Imagem de um logotipo.](../images/office15-app-create-wd-app-using-ooxml-fig03.png)

<span data-ttu-id="79fca-125">Você pode usar o mesmo método para inserir qualquer formato de imagem compatível com o Office.</span><span class="sxs-lookup"><span data-stu-id="79fca-125">You can use the same method for inserting any Office-supported image format.</span></span>

<span data-ttu-id="79fca-126">*Figura 4. Uma imagem formatada usando efeitos e estilos de imagem*</span><span class="sxs-lookup"><span data-stu-id="79fca-126">*Figure 4. An image formatted using picture styles and effects*</span></span>


![Imagem formatada no Word.](../images/office15-app-create-wd-app-using-ooxml-fig04.png)


<span data-ttu-id="79fca-128">A adição de efeitos e formatação de alta qualidade às imagens requer muito menos marcação do que você poderia esperar.</span><span class="sxs-lookup"><span data-stu-id="79fca-128">Adding high quality formatting and effects to your images requires much less markup than you might expect.</span></span>

<span data-ttu-id="79fca-129">*Figura 5. Um controle de conteúdo*</span><span class="sxs-lookup"><span data-stu-id="79fca-129">*Figure 5. A content control*</span></span>


![Texto em um controle de conteúdo vinculado.](../images/office15-app-create-wd-app-using-ooxml-fig05.png)

<span data-ttu-id="79fca-131">Você pode usar controles de conteúdo com o suplemento para adicionar conteúdo em um local especificado (associado) em vez de na seleção.</span><span class="sxs-lookup"><span data-stu-id="79fca-131">You can use content controls with your add-in to add content at a specified (bound) location rather than at the selection.</span></span>

<span data-ttu-id="79fca-132">*Figura 6. Uma caixa de texto com formatação do WordArt*</span><span class="sxs-lookup"><span data-stu-id="79fca-132">*Figure 6. A text box with WordArt formatting*</span></span>


![Texto formatado com efeitos de texto WordArt.](../images/office15-app-create-wd-app-using-ooxml-fig06.png)

<span data-ttu-id="79fca-134">Os efeitos de texto estão disponíveis no Word para o texto dentro de uma caixa de texto (como mostrado aqui) ou para o corpo do texto normal.</span><span class="sxs-lookup"><span data-stu-id="79fca-134">Text effects are available in Word for text inside a text box (as shown here) or for regular body text.</span></span>

<span data-ttu-id="79fca-135">*Figura 7. Uma forma*</span><span class="sxs-lookup"><span data-stu-id="79fca-135">*Figure 7. A shape*</span></span>


![Uma forma de desenho do Microsoft Office no Word.](../images/office15-app-create-wd-app-using-ooxml-fig07.png)

<span data-ttu-id="79fca-137">Você pode inserir formas de desenho internas ou personalizadas, com ou sem texto e efeitos de formatação.</span><span class="sxs-lookup"><span data-stu-id="79fca-137">You can insert built-in or custom drawing shapes, with or without text and formatting effects.</span></span>

<span data-ttu-id="79fca-138">*Figura 8. Uma tabela com formatação direta*</span><span class="sxs-lookup"><span data-stu-id="79fca-138">*Figure 8. A table with direct formatting*</span></span>


![Uma tabela formatada no Word.](../images/office15-app-create-wd-app-using-ooxml-fig08.png)

<span data-ttu-id="79fca-140">Você pode incluir formatação de texto, bordas, sombreamento, dimensionamento de células ou qualquer formatação de tabela que seja necessária.</span><span class="sxs-lookup"><span data-stu-id="79fca-140">You can include text formatting, borders, shading, cell sizing, or any table formatting you need.</span></span>

<span data-ttu-id="79fca-141">*Figura 9. Uma tabela formatada usando um estilo de tabela*</span><span class="sxs-lookup"><span data-stu-id="79fca-141">*Figure 9. A table formatted using a table style*</span></span>


![Uma tabela formatada no Word.](../images/office15-app-create-wd-app-using-ooxml-fig09.png)

<span data-ttu-id="79fca-143">Você pode usar estilos de tabela internos ou personalizados com a mesma facilidade com que usa um estilo de parágrafo para o texto.</span><span class="sxs-lookup"><span data-stu-id="79fca-143">You can use built-in or custom table styles just as easily as using a paragraph style for text.</span></span>

<span data-ttu-id="79fca-144">*Figura 10. Um diagrama do SmartArt*</span><span class="sxs-lookup"><span data-stu-id="79fca-144">*Figure 10. A SmartArt diagram*</span></span>


![Um diagrama SmartArt dinâmico no Word.](../images/office15-app-create-wd-app-using-ooxml-fig10.png)

<span data-ttu-id="79fca-146">O Microsoft Office oferece uma ampla variedade de layouts de diagrama do SmartArt (e você pode usar o Office Open XML para criar os seus próprios).</span><span class="sxs-lookup"><span data-stu-id="79fca-146">Microsoft Office offers a wide array of SmartArt diagram layouts (and you can use Office Open XML to create your own).</span></span>

<span data-ttu-id="79fca-147">*Figura 11. Um gráfico*</span><span class="sxs-lookup"><span data-stu-id="79fca-147">*Figure 11. A chart*</span></span>


![Um gráfico no Word.](../images/office15-app-create-wd-app-using-ooxml-fig11.png)

<span data-ttu-id="79fca-p105">Você pode inserir gráficos do Excel como gráficos dinâmicos em documentos do Word, o que também significa que você pode usá-los no seu suplemento do Word. Como você pode ver pelos exemplos anteriores, é possível usar a coerção do Office Open XML para inserir praticamente qualquer tipo de conteúdo que um usuário pode inserir em seu próprio documento. Há duas maneiras simples de obter a marcação do Office Open XML necessária. Adicionar conteúdo avançado a um documento do Word em branco e salvar o arquivo no formato de Documento XML do Word ou usar um suplemento de teste com o método [getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) para obter a marcação. As duas abordagens fornecem basicamente o mesmo resultado.</span><span class="sxs-lookup"><span data-stu-id="79fca-p105">You can insert Excel charts as live charts in Word documents, which also means you can use them in your add-in for Word. As you can see by the preceding examples, you can use Office Open XML coercion to insert essentially any type of content that a user can insert into their own document. There are two simple ways to get theOffice Open XML markup you need. Either add your rich content to an otherwise blank Word document and then save the file in Word XML Document format or use a test add-in with the [getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) method to grab the markup. Both approaches provide essentially the same result.</span></span>


> [!NOTE]
> <span data-ttu-id="79fca-p106">Um documento do Office Open XML é realmente um pacote compactado de arquivos que representa o conteúdo do documento. Salvar o arquivo no formato de Documento XML do Word lhe fornece todo o pacote do Office Open XML compactado em um arquivo XML, que também é o que você obtém ao usar **getSelectedDataAsync** para recuperar a marcação XML do Office Open XML.</span><span class="sxs-lookup"><span data-stu-id="79fca-p106">An Office Open XML document is actually a compressed package of files that represent the document contents. Saving the file in the Word XML Document format gives you the entireOffice Open XML package flattened into one XML file, which is also what you get when using  **getSelectedDataAsync** to retrieve the Office Open XML markup.</span></span>

<span data-ttu-id="79fca-p107">Se você salvar o arquivo em um formato XML do Word, observe que há duas opções na lista Salvar como Tipo na caixa de diálogo Salvar como para arquivos no formato .xml. Certifique-se de escolher **Documento XML do Word** e não a opção do Word 2003. Baixe o código de exemplo nomeado [Word-Add-in-Get-Set-EditOpen-XML](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML), que pode ser usado como uma ferramenta para recuperar e testar sua marcação. Então é só isso que preciso fazer? Bem, não exatamente. Sim, para muitos cenários, você poderia usar todo o resultado compactado do Office Open XML que obtém com um dos métodos anteriores, e ele funcionaria. A boa notícia é que você provavelmente não precisa da maioria dessa marcação. Se você é um dos muitos desenvolvedores de suplementos que estão vendo a marcação do Office Open XML pela primeira vez, tentar entender a grande quantidade de marcação obtida até para o conteúdo mais simples pode parecer assustador, mas não precisa ser assim. Neste tópico, usaremos alguns cenários comuns que obtivemos da comunidade de desenvolvedores de Suplementos do Office para mostrar técnicas que simplificam o Office Open XML para uso em suplementos. Exploraremos a marcação para alguns tipos de conteúdo mostrados anteriormente, além das informações necessárias para minimizar a carga do Office Open XML. Também examinaremos o código necessário para inserir conteúdo avançado em um documento na seleção ativa e a maneira de usar o Office Open XML com o objeto de associação para adicionar ou substituir conteúdo em locais específicos.</span><span class="sxs-lookup"><span data-stu-id="79fca-p107">If you save the file to an XML format from Word, note that there are two options under the Save as Type list in the Save As dialog box for .xml format files. Be sure to choose  **Word XML Document** and not the Word 2003 option. Download the code sample named [Word-Add-in-Get-Set-EditOpen-XML](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML), which you can use as a tool to retrieve and test your markup. So is that all there is to it? Well, not quite. Yes, for many scenarios, you could use the full, flattened Office Open XML result you see with either of the preceding methods and it would work. The good news is that you probably don't need most of that markup. If you're one of the many add-in developers seeing Office Open XML markup for the first time, trying to make sense of the massive amount of markup you get for the simplest piece of content might seem overwhelming, but it doesn't have to be. In this topic, we'll use some common scenarios we've been hearing from the Office Add-ins developer community to show you techniques for simplifying Office Open XML for use in your add-in. We'll explore the markup for some types of content shown earlier along with the information you need for minimizing the Office Open XML payload. We'll also look at the code you need for inserting rich content into a document at the active selection and how to use Office Open XML with the bindings object to add or replace content at specified locations.</span></span>

## <a name="exploring-the-office-open-xml-document-package"></a><span data-ttu-id="79fca-167">Explorar o pacote de documento do Office Open XML</span><span class="sxs-lookup"><span data-stu-id="79fca-167">Exploring the Office Open XML document package</span></span>


<span data-ttu-id="79fca-p108">Ao usar [getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) para recuperar o Office Open XML para uma seleção de conteúdo (ou ao salvar o documento no formato de Documento XML do Word), o que você obtém não é apenas a marcação que descreve o conteúdo selecionado, é um documento inteiro com várias opções e configurações das quais você certamente não necessita. De fato, se você usar esse método com um documento que contenha um suplemento de painel de tarefas, a marcação obtida incluirá até mesmo o painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="79fca-p108">When you use [getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) to retrieve the Office Open XML for a selection of content (or when you save the document in Word XML Document format), what you're getting is not just the markup that describes your selected content; it's an entire document with many options and settings that you almost certainly don't need. In fact, if you use that method from a document that contains a task pane add-in, the markup you get even includes your task pane.</span></span>

<span data-ttu-id="79fca-170">Até mesmo um pacote de documento simples do Word inclui partes para propriedades de documentos, estilos, tema (configurações de formatação), configurações da Web, fontes e muito mais, além de partes para o conteúdo real.</span><span class="sxs-lookup"><span data-stu-id="79fca-170">Even a simple Word document package includes parts for document properties, styles, theme (formatting settings), web settings, fonts, and then some, in addition to parts for the actual content.</span></span>

<span data-ttu-id="79fca-p109">Por exemplo, digamos que você queira inserir apenas um parágrafo de texto com formatação direta, conforme mostrado anteriormente na Figura 1. Ao usar o Office Open XML para o texto formatado com **getSelectedDataAsync**, você vê uma grande quantidade de marcação. A marcação inclui um elemento de pacote que representa um documento inteiro, que contém várias partes (comumente conhecidas como partes do documento ou, no Office Open XML, partes do pacote), como pode ver listado na Figura 13. Cada parte representa um arquivo separado dentro do pacote.</span><span class="sxs-lookup"><span data-stu-id="79fca-p109">For example, say that you want to insert just a paragraph of text with direct formatting, as shown earlier in Figure 1. When you grab the Office Open XML for the formatted text using  **getSelectedDataAsync**, you see a large amount of markup. That markup includes a package element that represents an entire document, which contains several parts (commonly referred to as document parts or, in the Office Open XML, as package parts), as you see listed in Figure 13. Each part represents a separate file within the package.</span></span>


> [!TIP]
> <span data-ttu-id="79fca-p110">Você pode editar a marcação do Office Open XML em um editor de texto como o Bloco de Notas. Se abri-lo no Visual Studio, pode usar **Editar > Avançado > Formatar Documento** (Ctrl+K, Ctrl+D) para formatar o pacote, facilitando a edição. Em seguida, você pode recolher ou expandir partes de um documento ou seções delas, conforme mostrado na Figura 12, para examinar e editar mais facilmente o conteúdo do pacote do Office Open XML. Cada parte do documento começa com uma marca **pkg:part**.</span><span class="sxs-lookup"><span data-stu-id="79fca-p110">You can edit Office Open XML markup in a text editor like Notepad. If you open it in Visual Studio, you can use  **Edit >Advanced > Format Document** (Ctrl+K, Ctrl+D) to format the package for easier editing. Then you can collapse or expand document parts or sections of them, as shown in Figure 12, to more easily review and edit the content of the Office Open XML package. Each document part begins with a **pkg:part** tag.</span></span>


<span data-ttu-id="79fca-179">*Figura 12. Recolher e expandir partes do pacote para facilitar a edição no Visual Studio*</span><span class="sxs-lookup"><span data-stu-id="79fca-179">*Figure 12. Collapse and expand package parts for easier editing in Visual Studio*</span></span>

![Trecho de código do Office Open XML de uma parte de pacote.](../images/office15-app-create-wd-app-using-ooxml-fig12.png)

<span data-ttu-id="79fca-181">*Figura 13. As partes incluídas em um pacote de documento básico do Office Open XML do Word*</span><span class="sxs-lookup"><span data-stu-id="79fca-181">*Figure 13. The parts included in a basic Word Office Open XML document package*</span></span>

![Trecho de código do Office Open XML de uma parte de pacote.](../images/office15-app-create-wd-app-using-ooxml-fig13.png)

<span data-ttu-id="79fca-183">Com toda essa marcação, você poderá se surpreender ao descobrir que os únicos elementos realmente necessários para inserir o exemplo de texto formatado são pedaços da parte .rels e a parte document.xml.</span><span class="sxs-lookup"><span data-stu-id="79fca-183">With all that markup, you might be surprised to discover that the only elements you actually need to insert the formatted text example are pieces of the .rels part and the document.xml part.</span></span>


> [!NOTE]
> <span data-ttu-id="79fca-p111">As duas linhas de marcação acima da marca do pacote (as declarações de XML para a versão e a ID do programa do Office) são pressupostas quando você usa o tipo de coerção do Office Open XML, assim, não é preciso incluí-las. Mantenha-as se você quiser abrir a marcação editada como um documento do Word para testá-la.</span><span class="sxs-lookup"><span data-stu-id="79fca-p111">The two lines of markup above the package tag (the XML declarations for version and Office program ID) are assumed when you use the Office Open XML coercion type, so you don't need to include them. Keep them if you want to open your edited markup as a Word document to test it.</span></span>

<span data-ttu-id="79fca-p112">Vários dos outros tipos de conteúdo mostrados no início deste tópico também exigem partes adicionais (além daquelas mostradas na Figura 13), e vamos abordá-los mais adiante neste tópico. Enquanto isso, como você verá a maioria das partes mostradas na Figura 13 na marcação de qualquer pacote de documento do Word, aqui está um resumo rápido do que cada uma das partes faz e quando você precisa delas:</span><span class="sxs-lookup"><span data-stu-id="79fca-p112">Several of the other types of content shown at the start of this topic require additional parts as well (beyond those shown in Figure 13), and we'll address those later in this topic. Meanwhile, since you'll see most of the parts shown in Figure 13 in the markup for any Word document package, here's a quick summary of what each of these parts is for and when you need it:</span></span>


- <span data-ttu-id="79fca-p113">Dentro da marca de pacote, a primeira parte é o arquivo .rels, que define as relações entre as partes de nível superior do pacote (elas normalmente são as propriedades do documento, a miniatura, se houver, e o corpo do documento principal). Sempre é necessário algum conteúdo nessa parte na marcação, pois você precisa definir a relação entre a parte do documento principal (em que o conteúdo reside) e o pacote de documento.</span><span class="sxs-lookup"><span data-stu-id="79fca-p113">Inside the package tag, the first part is the .rels file, which defines relationships between the top-level parts of the package (these are typically the document properties, thumbnail (if any), and main document body). Some of the content in this part is always required in your markup because you need to define the relationship of the main document part (where your content resides) to the document package.</span></span>

- <span data-ttu-id="79fca-190">A parte document.xml.rels define as relações para as partes adicionais necessárias para a parte document.xml (corpo principal), se houver.</span><span class="sxs-lookup"><span data-stu-id="79fca-190">The document.xml.rels part defines relationships for additional parts required by the document.xml (main body) part, if any.</span></span>


   > [!IMPORTANT]
   > <span data-ttu-id="79fca-p114">Os arquivos .rels no pacote (como .rels de nível superior, document.xml.rels e outros que você pode ver para tipos específicos de conteúdo) são uma ferramenta extremamente importante que você pode usar como guia para ajudá-lo a editar rapidamente o pacote do Office Open XML. Para saber mais sobre como fazer isso, confira [Criar sua própria marcação: práticas recomendadas](#creating-your-own-markup-best-practices) mais adiante neste tópico.</span><span class="sxs-lookup"><span data-stu-id="79fca-p114">The .rels files in your package (such as the top-level .rels, document.xml.rels, and others you may see for specific types of content) are an extremely important tool that you can use as a guide for helping you quickly edit down your Office Open XML package. To learn more about how to do this, see [Creating your own markup: best practices](#creating-your-own-markup-best-practices) later in this topic.</span></span>



- <span data-ttu-id="79fca-p115">A parte document.xml é o conteúdo no corpo principal do documento. Você precisa de elementos dessa parte, claro, pois é onde o conteúdo aparece. Porém, você não precisa de tudo o que vê nessa parte. Examinaremos isso em mais detalhes posteriormente.</span><span class="sxs-lookup"><span data-stu-id="79fca-p115">The document.xml part is the content in the main body of the document. You need elements of this part, of course, since that's where your content appears. But, you don't need everything you see in this part. We'll look at that in more detail later.</span></span>

- <span data-ttu-id="79fca-p116">Muitas partes são automaticamente ignoradas pelos métodos Set ao se inserir conteúdo em um documento usando a coerção do Office Open XML, assim, você pode removê-las. Isso inclui o arquivo theme1.xml (o tema de formatação do documento), as partes de propriedades do documento (núcleo, suplemento e miniatura) e arquivos de configurações (incluindo settings, webSettings e fontTable).</span><span class="sxs-lookup"><span data-stu-id="79fca-p116">Many parts are automatically ignored by the Set methods when inserting content into a document using Office Open XML coercion, so you might as well remove them. These include the theme1.xml file (the document's formatting theme), the document properties parts (core, add-in, and thumbnail), and setting files (including settings, webSettings, and fontTable).</span></span>

- <span data-ttu-id="79fca-p117">No exemplo da Figura 1, a formatação de texto é aplicada diretamente (ou seja, cada configuração de fonte e de formatação de parágrafo é aplicada individualmente). Contudo, se você usar um estilo (por exemplo, se desejar que o texto assuma automaticamente a formatação do estilo Título 1 no documento de destino) como mostrado anteriormente na Figura 2, precisará da parte styles.xml, bem como de uma definição de relacionamento para ele. Para saber mais, confira a seção do tópico [Adicionar objetos que usam partes adicionais do Office Open XML](#adding-objects-that-use-additional-office-open-xml-parts).</span><span class="sxs-lookup"><span data-stu-id="79fca-p117">In the Figure 1 example, text formatting is directly applied (that is, each font and paragraph formatting setting applied individually). But, if you use a style (such as if you want your text to automatically take on the formatting of the Heading 1 style in the destination document) as shown earlier in Figure 2, then you would need part of the styles.xml part as well as a relationship definition for it. For more information, see the topic section [Adding objects that use additional Office Open XML parts](#adding-objects-that-use-additional-office-open-xml-parts).</span></span>


## <a name="inserting-document-content-at-the-selection"></a><span data-ttu-id="79fca-202">Inserir conteúdo de documento na seleção</span><span class="sxs-lookup"><span data-stu-id="79fca-202">Inserting document content at the selection</span></span>


<span data-ttu-id="79fca-203">Vamos examinar a marcação mínima do Office Open XML necessária para o exemplo de texto formatado mostrado na Figura 1 e o JavaScript necessário para inseri-la na seleção ativa no documento.</span><span class="sxs-lookup"><span data-stu-id="79fca-203">Let's take a look at the minimal Office Open XML markup required for the formatted text example shown in Figure 1 and the JavaScript required for inserting it at the active selection in the document.</span></span>


### <a name="simplified-office-open-xml-markup"></a><span data-ttu-id="79fca-204">Marcação simplificada do Office Open XML</span><span class="sxs-lookup"><span data-stu-id="79fca-204">Simplified Office Open XML markup</span></span>

<span data-ttu-id="79fca-p118">Editamos o exemplo do Office Open XML mostrado aqui, conforme descrito na seção anterior, para deixar apenas as partes do documento obrigatórias e somente os elementos necessários em cada uma dessas partes. Vamos examinar como editar a marcação você mesmo (e explicar um pouco mais as partes restantes aqui) na próxima seção do tópico.</span><span class="sxs-lookup"><span data-stu-id="79fca-p118">We've edited the Office Open XML example shown here, as described in the preceding section, to leave just required document parts and only required elements within each of those parts. We'll walk through how to edit the markup yourself (and explain a bit more about the pieces that remain here) in the next section of the topic.</span></span>


```XML
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" >
        <w:body>
          <w:p>
            <w:pPr>
              <w:spacing w:before="360" w:after="0" w:line="480" w:lineRule="auto"/>
              <w:rPr>
                <w:color w:val="70AD47" w:themeColor="accent6"/>
                <w:sz w:val="28"/>
              </w:rPr>
            </w:pPr>
            <w:r>
              <w:rPr>
                <w:color w:val="70AD47" w:themeColor="accent6"/>
                <w:sz w:val="28"/>
              </w:rPr>
              <w:t>This text has formatting directly applied to achieve its font size, color, line spacing, and paragraph spacing.</w:t>
            </w:r>
          </w:p>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>
```


> [!NOTE]
> <span data-ttu-id="79fca-p119">Se você adicionar a marcação mostrada aqui a um arquivo XML com as marcas de declaração de XML para versão e mso-application na parte superior do arquivo (mostrado na Figura 13), você poderá abri-lo no Word como um documento do Word. Ou, sem essas marcas, ainda poderá abri-lo usando **Arquivo > Abrir** no Word. Você verá **Modo de Compatibilidade** na barra de título no Word, pois removeu as configurações que avisam ao Word que se trata de um documento. Como você está adicionando a marcação a um documento existente do Word, isso não afetará o conteúdo de forma alguma.</span><span class="sxs-lookup"><span data-stu-id="79fca-p119">If you add the markup shown here to an XML file along with the XML declaration tags for version and mso-application at the top of the file (shown in Figure 13), you can open it in Word as a Word document. Or, without those tags, you can still open it using  **File> Open** in Word. You'll see **Compatibility Mode** on the title bar in Word, because you removed the settings that tell Word this is a Word document. Since you're adding this markup to an existing Word document, that won't affect your content at all.</span></span>


### <a name="javascript-for-using-setselecteddataasync"></a><span data-ttu-id="79fca-211">JavaScript para usar setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="79fca-211">JavaScript for using setSelectedDataAsync</span></span>


<span data-ttu-id="79fca-212">Após salvar o Office Open XML anterior como um arquivo XML que pode ser acessado por meio de sua solução, você poderá usar a função a seguir para definir o conteúdo de texto formatado no documento usando a coerção do Office Open XML.</span><span class="sxs-lookup"><span data-stu-id="79fca-212">Once you save the preceding Office Open XML as an XML file that's accessible from your solution, you can use the following function to set the formatted text content in the document using Office Open XML coercion.</span></span> 

<span data-ttu-id="79fca-p120">Nessa função, observe que, exceto pela última linha, tudo é usado para acessar a marcação salva para uso na chamada de método [setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) no fim da função. **setSelectedDataASync** requer apenas que você especifique o conteúdo a ser inserido e o tipo de coerção.</span><span class="sxs-lookup"><span data-stu-id="79fca-p120">In this function, notice that all but the last line are used to get your saved markup for use in the [setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method call at the end of the function. **setSelectedDataASync** requires only that you specify the content to be inserted and the coercion type.</span></span>


> [!NOTE]
> <span data-ttu-id="79fca-p121">Substitua _yourXMLfilename_ pelo nome e pelo caminho do arquivo XML que você salvou na solução. Se não tiver certeza de onde incluir arquivos XML na solução ou como referenciá-los no código, confira o exemplo de código [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML) para obter exemplos disso e um exemplo operacional da marcação e do JavaScript mostrado aqui.</span><span class="sxs-lookup"><span data-stu-id="79fca-p121">Replace  _yourXMLfilename_ with the name and path of the XML file as you've saved it in your solution. If you're not sure where to include XML files in your solution or how to reference them in your code, see the [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML) code sample for examples of that and a working example of the markup and JavaScript shown here.</span></span>




```js
function writeContent() {
    var myOOXMLRequest = new XMLHttpRequest();
    var myXML;
    myOOXMLRequest.open('GET', 'yourXMLfilename', false);
    myOOXMLRequest.send();
    if (myOOXMLRequest.status === 200) {
        myXML = myOOXMLRequest.responseText;
    }
    Office.context.document.setSelectedDataAsync(myXML, { coercionType: 'ooxml' });
}
```


## <a name="creating-your-own-markup-best-practices"></a><span data-ttu-id="79fca-217">Criar sua própria marcação: práticas recomendadas</span><span class="sxs-lookup"><span data-stu-id="79fca-217">Creating your own markup: best practices</span></span>


<span data-ttu-id="79fca-218">Vamos examinar mais detalhadamente a marcação que deve ser inserida no exemplo de texto formatado anterior.</span><span class="sxs-lookup"><span data-stu-id="79fca-218">Let's take a closer look at the markup you need to insert the preceding formatted text example.</span></span>

<span data-ttu-id="79fca-p122">Para o exemplo, comece simplesmente excluindo todas as partes de documento do pacote, exceto .rels e document.xml. Em seguida, editaremos essas duas partes necessárias para simplificar tudo ainda mais.</span><span class="sxs-lookup"><span data-stu-id="79fca-p122">For this example, start by simply deleting all document parts from the package other than .rels and document.xml. Then, we'll edit those two required parts to simplify things further.</span></span>


> [!IMPORTANT]
> <span data-ttu-id="79fca-p123">Use as partes .rels como um mapa para avaliar rapidamente o que está incluído no pacote e determinar quais partes você pode excluir completamente (ou seja, as partes não relacionadas ou nem referenciadas pelo conteúdo). Lembre-se de que todas as partes do documento devem ter uma relação definida no pacote e as relações aparecem nos arquivos .rels. Assim, você deve ver todas elas listadas em .rels, em document.xml.rels ou em um arquivo .rels específico do conteúdo.</span><span class="sxs-lookup"><span data-stu-id="79fca-p123">Use the .rels parts as a map to quickly gauge what's included in the package and determine what parts you can delete completely (that is, any parts not related to or referenced by your content). Remember that every document part must have a relationship defined in the package and those relationships appear in the .rels files. So you should see all of them listed in either .rels, document.xml.rels, or a content-specific .rels file.</span></span>

<span data-ttu-id="79fca-p124">A marcação a seguir mostra a parte .rels necessária antes da edição. Como estamos excluindo o suplemento, partes de propriedade do documento principal e a parte de miniatura, também precisamos excluir essas relações de .rels. Observe que isso deixará somente a relação (com a ID de relação "rID1" no exemplo a seguir) para document.xml.</span><span class="sxs-lookup"><span data-stu-id="79fca-p124">The following markup shows the required .rels part before editing. Since we're deleting the add-in and core document property parts, and the thumbnail part, we need to delete those relationships from .rels as well. Notice that this will leave only the relationship (with the relationship ID "rID1" in the following example) for document.xml.</span></span>




```XML
<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512">
  <pkg:xmlData>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
      <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail" Target="docProps/thumbnail.emf"/>
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
    </Relationships>
  </pkg:xmlData>
</pkg:part>
```


> [!IMPORTANT]
> <span data-ttu-id="79fca-p125">Remova as relações (ou seja, a marca **Relationship**) de todas as partes que você remover completamente do pacote. Incluir uma parte sem uma relação correspondente ou excluir uma parte e deixar sua relação no pacote resultará em um erro.</span><span class="sxs-lookup"><span data-stu-id="79fca-p125">Remove the relationships (that is, the **Relationship** tag) for any parts that you completely remove from the package. Including a part without a corresponding relationship, or excluding a part and leaving its relationship in the package, will result in an error.</span></span>

<span data-ttu-id="79fca-229">A marcação a seguir mostra a parte document.xml, que inclui o conteúdo de texto formatado de exemplo antes da edição.</span><span class="sxs-lookup"><span data-stu-id="79fca-229">The following markup shows the document.xml part, which includes our sample formatted text content before editing.</span></span>

```XML
<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document mc:Ignorable="w14 w15 wp14" xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
        <w:body>
          <w:p>
            <w:pPr>
              <w:spacing w:before="360" w:after="0" w:line="480" w:lineRule="auto"/>
              <w:rPr>
                <w:color w:val="70AD47" w:themeColor="accent6"/>
                <w:sz w:val="28"/>
              </w:rPr>
            </w:pPr>
            <w:r>
              <w:rPr>
                <w:color w:val="70AD47" w:themeColor="accent6"/>
                <w:sz w:val="28"/>
              </w:rPr>
              <w:t>This text has formatting directly applied to achieve its font size, color, line spacing, and paragraph spacing.</w:t>
            </w:r>
            <w:bookmarkStart w:id="0" w:name="_GoBack"/>
            <w:bookmarkEnd w:id="0"/>
          </w:p>
          <w:p/>
          <w:sectPr>
            <w:pgSz w:w="12240" w:h="15840"/>
            <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
            <w:cols w:space="720"/>
          </w:sectPr>
        </w:body>
      </w:document>
    </pkg:xmlData>
</pkg:part>
```

<span data-ttu-id="79fca-p126">Como document.xml é a parte do documento principal em que você coloca o conteúdo, vamos dar uma olhada rápida nessa parte. (A Figura 14, exibida após a lista, fornece uma referência visual para mostrar como parte do conteúdo principal e das marcas de formatação explicadas aqui se relacionam ao que você vê em um documento do Word.)</span><span class="sxs-lookup"><span data-stu-id="79fca-p126">Since document.xml is the primary document part where you place your content, let's take a quick walk through that part. (Figure 14, which follows this list, provides a visual reference to show how some of the core content and formatting tags explained here relate to what you see in a Word document.)</span></span>


- <span data-ttu-id="79fca-p127">A marca de abertura **w:document** inclui várias listagens de namespaces (**xmlns**). Muitos desses namespaces referem-se a tipos específicos de conteúdo, e você só precisa deles caso sejam relevantes para o conteúdo.</span><span class="sxs-lookup"><span data-stu-id="79fca-p127">The opening **w:document** tag includes several namespace ( **xmlns** ) listings. Many of those namespaces refer to specific types of content and you only need them if they're relevant to your content.</span></span>

    <span data-ttu-id="79fca-p128">O prefixo para as marcas em uma parte do documento remete aos namespaces. Neste exemplo, o único prefixo usado nas marcas em todo o document.xml é **w:**, portanto o único namespace que precisamos deixar na marca de abertura **w:document** é **xmlns:w**.</span><span class="sxs-lookup"><span data-stu-id="79fca-p128">Notice that the prefix for the tags throughout a document part refers back to the namespaces. In this example, the only prefix used in the tags throughout the document.xml part is  **w:**, so the only namespace that we need to leave in the opening **w:document** tag is **xmlns:w**.</span></span>


> [!TIP]
> <span data-ttu-id="79fca-p129">Se você estiver editando a marcação no Visual Studio, após excluir namespaces em qualquer parte, examine todas as marcas dessa parte. Se tiver removido um namespace necessário para a marcação, você verá um pequeno sublinhado ondulado vermelho no prefixo relevante das marcas afetadas. Se remover o namespace **xmlns:mc**, você também deverá remover o atributo **mc:Ignorable** que precede as listagens de namespace.</span><span class="sxs-lookup"><span data-stu-id="79fca-p129">If you're editing your markup in Visual Studio, after you delete namespaces in any part, look through all tags of that part. If you've removed a namespace that's required for your markup, you'll see a red squiggly underline on the relevant prefix for affected tags. If you remove the **xmlns:mc** namespace, you must also remove the **mc:Ignorable** attribute that precedes the namespace listings.</span></span>


- <span data-ttu-id="79fca-239">Dentro da marca de abertura do corpo, você verá uma marca de parágrafo (**w:p**), que inclui o conteúdo para este exemplo.</span><span class="sxs-lookup"><span data-stu-id="79fca-239">Inside the opening body tag, you see a paragraph tag ( **w:p** ), which includes our sample content for this example.</span></span>

- <span data-ttu-id="79fca-p130">A marca **w:pPr** inclui propriedades para formatação de parágrafo aplicada diretamente, como um espaço antes ou depois do parágrafo, o alinhamento do parágrafo ou os recuos. (A formatação direta refere-se aos atributos que você aplica individualmente ao conteúdo, não como parte de um estilo.) Essa marca também inclui formatação de fonte direta que é aplicada a todo o parágrafo, em uma marca aninhada **w:rPr** (propriedades de execução), que contém a cor da fonte e o tamanho definido no exemplo.</span><span class="sxs-lookup"><span data-stu-id="79fca-p130">The **w:pPr** tag includes properties for directly-applied paragraph formatting, such as space before or after the paragraph, paragraph alignment, or indents. (Direct formatting refers to attributes that you apply individually to content rather than as part of a style.) This tag also includes direct font formatting that's applied to the entire paragraph, in a nested **w:rPr** (run properties) tag, which contains the font color and size set in our sample.</span></span>


   > [!NOTE]
   > <span data-ttu-id="79fca-p131">Talvez você perceba que tamanhos de fonte e outras configurações de formatação na marcação do Word do Office Open XML parecem ter o dobro do tamanho real. Isso ocorre porque o espaçamento de parágrafo e linha, bem como algumas propriedades de formatação de seção mostradas na marcação anterior, são especificados em twips (um vigésimo de um ponto). Dependendo dos tipos de conteúdo com os quais trabalha no Office Open XML, você pode ver várias unidades de medida adicionais, incluindo Unidades Métricas em Inglês (914.400 EMUs para uma polegada), que são usadas para alguns valores de Arte do Office (drawingML) e 100.000 vezes o valor real, que é usado em drawingML e na marcação do PowerPoint. O PowerPoint também expressa alguns valores como 100 vezes o valor real, e o Excel comumente usa os valores reais.</span><span class="sxs-lookup"><span data-stu-id="79fca-p131">You might notice that font sizes and some other formatting settings in Word Office Open XML markup look like they're double the actual size. That's because paragraph and line spacing, as well some section formatting properties shown in the preceding markup, are specified in twips (one-twentieth of a point). Depending on the types of content you work with in Office Open XML, you may see several additional units of measure, including English Metric Units (914,400 EMUs to an inch), which are used for some Office Art (drawingML) values and 100,000 times actual value, which is used in both drawingML and PowerPoint markup. PowerPoint also expresses some values as 100 times actual and Excel commonly uses actual values.</span></span>


- <span data-ttu-id="79fca-p132">Em um parágrafo, qualquer conteúdo com propriedades semelhantes é incluído em uma execução (**w:r**), como é o caso do texto de exemplo. Sempre que há uma alteração no tipo de conteúdo ou formatação, uma nova execução é iniciada. (Ou seja, se apenas uma palavra no texto de exemplo estivesse em negrito, ela seria separada em sua própria execução.) Neste exemplo, o conteúdo inclui apenas o texto de uma execução.</span><span class="sxs-lookup"><span data-stu-id="79fca-p132">Within a paragraph, any content with like properties is included in a run ( **w:r** ), such as is the case with the sample text. Each time there's a change in formatting or content type, a new run starts. (That is, if just one word in the sample text was bold, it would be separated into its own run.) In this example, the content includes just the one text run.</span></span>

    <span data-ttu-id="79fca-249">Como a formatação incluída neste exemplo é a formatação da fonte (ou seja, a formatação que pode ser aplicada a apenas um caractere), ela também aparece nas propriedades para a execução individual.</span><span class="sxs-lookup"><span data-stu-id="79fca-249">Notice that, because the formatting included in this sample is font formatting (that is, formatting that can be applied to as little as one character), it also appears in the properties for the individual run.</span></span>

- <span data-ttu-id="79fca-p133">Observe também as marcas para o indicador oculto "_GoBack" (**w:bookmarkStart** e **w:bookmarkEnd**), que aparecem nos documentos do Word por padrão. Você sempre pode excluir as marcas de início e de término do indicador GoBack da marcação.</span><span class="sxs-lookup"><span data-stu-id="79fca-p133">Also notice the tags for the hidden "_GoBack" bookmark (**w:bookmarkStart** and **w:bookmarkEnd** ), which appear in Word documents by default. You can always delete the start and end tags for the GoBack bookmark from your markup.</span></span>

- <span data-ttu-id="79fca-p134">A última parte do corpo do documento é a marca **w:sectPr**, ou propriedades de seção. Essa marca inclui configurações como margens e orientação da página. O conteúdo que você inserir usando **setSelectedDataAsync** adotará as propriedades da seção ativa no documento de destino por padrão. Portanto, a menos que o conteúdo inclua uma quebra de seção (nesse caso, haverá mais de uma marca **w:sectPr**), você pode excluir essa marca.</span><span class="sxs-lookup"><span data-stu-id="79fca-p134">The last piece of the document body is the **w:sectPr** tag, or section properties. This tag includes settings such as margins and page orientation. The content you insert using **setSelectedDataAsync** will take on the active section properties in the destination document by default. So, unless your content includes a section break (in which case you'll see more than one **w:sectPr** tag), you can delete this tag.</span></span>


<span data-ttu-id="79fca-256">*Figura 14. Como marcas comuns em document.xml estão relacionadas ao conteúdo e ao layout de um documento do Word*</span><span class="sxs-lookup"><span data-stu-id="79fca-256">*Figure 14. How common tags in document.xml relate to the content and layout of a Word document*</span></span>

![Elementos do Office Open XML em um documento do Word.](../images/office15-app-create-wd-app-using-ooxml-fig14.png)

> [!TIP]
> <span data-ttu-id="79fca-p135">Na marcação que você criar, talvez haja outro atributo em várias marcas que inclui os caracteres **w:rsid**, que você não vê nos exemplos usados neste tópico. Esses são identificadores de revisão. Eles são usados no Word para o recurso Combinar Documentos e estão ativados por padrão. Você nunca precisará deles na marcação que está inserindo com o suplemento, e desativá-los torna a marcação bem mais limpa. Você pode facilmente remover marcas RSID existentes ou desabilitar o recurso (conforme descrito no procedimento a seguir) para que eles não sejam adicionados à marcação para o novo conteúdo.</span><span class="sxs-lookup"><span data-stu-id="79fca-p135">In markup you create, you might see another attribute in several tags that includes the characters **w:rsid**, which you don't see in the examples used in this topic. These are revision identifiers. They're used in Word for the Combine Documents feature and they're on by default. You'll never need them in markup you're inserting with your add-in and turning them off makes for much cleaner markup. You can easily remove existing RSID tags or disable the feature (as described in the following procedure) so that they're not added to your markup for new content.</span></span>

<span data-ttu-id="79fca-263">Lembre-se de que se você usar os recursos de coautoria no Word (como a capacidade de editar simultaneamente documentos com outras pessoas), você deve ativar o recurso novamente quando tiver terminado de gerar a marcação para seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="79fca-263">Be aware that if you use the co-authoring capabilities in Word (such as the ability to simultaneously edit documents with others), you should enable the feature again when finished generating the markup for your add-in.</span></span>

<span data-ttu-id="79fca-264">Para desativar atributos RSID no Word para documentos que você criar no futuro, faça o seguinte:</span><span class="sxs-lookup"><span data-stu-id="79fca-264">To turn off RSID attributes in Word for documents you create going forward, do the following:</span></span> 

1. <span data-ttu-id="79fca-265">No Word, escolha a guia **Arquivo** e escolha **Opções**.</span><span class="sxs-lookup"><span data-stu-id="79fca-265">In Word, choose **File** and then choose **Options**.</span></span>
2. <span data-ttu-id="79fca-266">Na caixa de diálogo Opções do Word, escolha **Central de Confiabilidade** e escolha **Configurações da Central de Confiabilidade**.</span><span class="sxs-lookup"><span data-stu-id="79fca-266">In the Word Options dialog box, choose **Trust Center** and then choose **Trust Center Settings**.</span></span>
3. <span data-ttu-id="79fca-267">Na caixa de diálogo Central de Confiabilidade, escolha **Opções de privacidade** e desative a configuração **Armazenar número aleatório para melhorar a precisão da combinação**.</span><span class="sxs-lookup"><span data-stu-id="79fca-267">In the Trust Center dialog box, choose **Privacy Options** and then disable the setting **Store random numbers to improve Combine accuracy**.</span></span>

<span data-ttu-id="79fca-268">Para remover marcas RSID de um documento existente, tente o seguinte atalho com o documento aberto no Office Open XML:</span><span class="sxs-lookup"><span data-stu-id="79fca-268">To remove RSID tags from an existing document, try the following shortcut with the document open in Office Open XML:</span></span>


1. <span data-ttu-id="79fca-269">Com o ponto de inserção no corpo do documento principal, pressione **Ctrl+Home** para ir para a parte superior do documento.</span><span class="sxs-lookup"><span data-stu-id="79fca-269">With your insertion point in the main body of the document, press **Ctrl+Home** to go to the top of the document.</span></span>
2. <span data-ttu-id="79fca-p136">No teclado, pressione **Barra de espaços**, **Delete**, **Barra de espaços**. Em seguida, salve o documento.</span><span class="sxs-lookup"><span data-stu-id="79fca-p136">On the keyboard, press **Spacebar**, **Delete**, **Spacebar**. Then, save the document.</span></span>

<span data-ttu-id="79fca-272">Após remover a maior parte da marcação do pacote, resta a marcação mínima que precisa ser inserida para o exemplo, conforme mostrado na seção anterior.</span><span class="sxs-lookup"><span data-stu-id="79fca-272">After removing the majority of the markup from this package, we're left with the minimal markup that needs to be inserted for the sample, as shown in the preceding section.</span></span>


## <a name="using-the-same-office-open-xml-structure-for-different-content-types"></a><span data-ttu-id="79fca-273">Usar a mesma estrutura do Office Open XML para diferentes tipos de conteúdo</span><span class="sxs-lookup"><span data-stu-id="79fca-273">Using the same Office Open XML structure for different content types</span></span>


<span data-ttu-id="79fca-p137">Vários tipos de conteúdo avançado exigem somente os componentes .rels e document.xml mostrados no exemplo anterior, incluindo controles de conteúdo, formas de desenho e caixas de texto do Office e tabelas (a menos que um estilo seja aplicado à tabela). De fato, você pode reutilizar as mesmas partes de pacote editadas e trocar apenas o conteúdo de **body** em document.xml para a marcação do conteúdo.</span><span class="sxs-lookup"><span data-stu-id="79fca-p137">Several types of rich content require only the .rels and document.xml components shown in the preceding example, including content controls, Office drawing shapes and text boxes, and tables (unless a style is applied to the table). In fact, you can reuse the same edited package parts and swap out just the **body** content in document.xml for the markup of your content.</span></span>

<span data-ttu-id="79fca-276">Para verificar a marcação do Office Open XML para os exemplos de cada um dos tipos de conteúdo mostrados anteriormente nas Figuras 5 a 8, explore o exemplo de código [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML) referenciado na seção Visão geral.</span><span class="sxs-lookup"><span data-stu-id="79fca-276">To check out the Office Open XML markup for the examples of each of these content types shown earlier in Figures 5 through 8, explore the [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML) code sample referenced in the overview section.</span></span>

<span data-ttu-id="79fca-277">Antes de continuarmos, vamos dar uma olhada nas diferenças relevantes para alguns desses tipos de conteúdo e como trocar as partes de que você necessita.</span><span class="sxs-lookup"><span data-stu-id="79fca-277">Before we move on, let's take a look at differences to note for a couple of these content types and how to swap out the pieces you need.</span></span>


### <a name="understanding-drawingml-markup-office-graphics-in-word-what-are-fallbacks"></a><span data-ttu-id="79fca-278">Compreender a marcação de drawingML (elementos gráficos do Office) no Word: O que são fallbacks?</span><span class="sxs-lookup"><span data-stu-id="79fca-278">Understanding drawingML markup (Office graphics) in Word: What are fallbacks?</span></span>

<span data-ttu-id="79fca-p138">Se a marcação da forma ou da caixa de texto parece muito mais complexa do que o esperado, há um motivo para isso. Com o lançamento do Office 2007, houve a introdução dos Formatos do Office Open XML e de um novo mecanismo de elementos gráficos do Office que o PowerPoint e o Excel adotaram plenamente. Na versão 2007, o Word só incorporou parte desse mecanismo de elementos gráficos, adotando o mecanismo de elementos gráficos atualizado do Excel, elementos gráficos SmartArt e ferramentas de imagem avançadas. Para formas e caixas de texto, o Word 2007 continua a usar objetos de desenho herdados (VML). Na versão 2010, o Word lançou etapas adicionais com o mecanismo de elementos gráficos para incorporar formas e ferramentas de desenho atualizadas.</span><span class="sxs-lookup"><span data-stu-id="79fca-p138">If the markup for your shape or text box looks far more complex than you would expect, there is a reason for it. With the release of Office 2007, we saw the introduction of the Office Open XML Formats as well as the introduction of a new Office graphics engine that PowerPoint and Excel fully adopted. In the 2007 release, Word only incorporated part of that graphics engine, adopting the updated Excel charting engine, SmartArt graphics, and advanced picture tools. For shapes and text boxes, Word 2007 continued to use legacy drawing objects (VML). It was in the 2010 release that Word took the additional steps with the graphics engine to incorporate updated shapes and drawing tools.</span></span>

<span data-ttu-id="79fca-284">Portanto, para dar suporte a formas e caixas de texto em documentos do Word no Formato do Office Open XML quando abertos no Word 2007, as formas (incluindo caixas de texto) exigem marcação VML de fallback.</span><span class="sxs-lookup"><span data-stu-id="79fca-284">So, to support shapes and text boxes in Office Open XML Format Word documents when opened in Word 2007, shapes (including text boxes) require fallback VML markup.</span></span>

<span data-ttu-id="79fca-p139">Normalmente, como você pode ver nos exemplos de forma e caixa de texto incluídos no exemplo de código [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML), a marcação da reparação pode ser removida. O Word adiciona automaticamente a marcação de reparação ausente às formas quando um documento é salvo. No entanto, se você prefere manter a marcação de reparação para garantir o suporte a todos os cenários de usuário, não há problema em mantê-la.</span><span class="sxs-lookup"><span data-stu-id="79fca-p139">Typically, as you see for the shape and text box examples included in the [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML) code sample, the fallback markup can be removed. Word automatically adds missing fallback markup to shapes when a document is saved. However, if you prefer to keep the fallback markup to ensure that you're supporting all user scenarios, there's no harm in retaining it.</span></span>

<span data-ttu-id="79fca-p140">Se houver objetos de desenho agrupados incluídos no conteúdo, você verá marcação adicional (e aparentemente repetitiva), mas isso deve ser mantido. Partes da marcação para formas de desenho são duplicadas quando o objeto é incluído em um grupo.</span><span class="sxs-lookup"><span data-stu-id="79fca-p140">If you have grouped drawing objects included in your content, you'll see additional (and apparently repetitive) markup, but this must be retained. Portions of the markup for drawing shapes are duplicated when the object is included in a group.</span></span>


> [!IMPORTANT]
> <span data-ttu-id="79fca-p141">Ao trabalhar com caixas de texto e formas de desenho, verifique os namespaces cuidadosamente antes de removê-los de document.xml. (Ou então, se você estiver reutilizando marcação de outro tipo de objeto, adicione novamente quaisquer namespaces necessários que tenham sido removidos anteriormente de document.xml.) Uma parte substancial dos namespaces incluídos por padrão em document.xml está presente devido a requisitos de objeto de desenho.</span><span class="sxs-lookup"><span data-stu-id="79fca-p141">When working with text boxes and drawing shapes, be sure to check namespaces carefully before removing them from document.xml. (Or, if you're reusing markup from another object type, be sure to add back any required namespaces you might have previously removed from document.xml.) A substantial portion of the namespaces included by default in document.xml are there for drawing object requirements.</span></span>


#### <a name="about-graphic-positioning"></a><span data-ttu-id="79fca-292">Sobre o posicionamento de gráficos</span><span class="sxs-lookup"><span data-stu-id="79fca-292">About graphic positioning</span></span>

<span data-ttu-id="79fca-p142">Nos exemplos de código [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML) e [Word-Add-in-Get-Set-EditOpen-XML](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML), a caixa de texto e a forma são configuradas usando diferentes tipos de configurações de posicionamento e disposição de texto. (Lembre-se também de que os exemplos de imagem nesses exemplos de código são configurados usando formatação embutida com texto, que posiciona um objeto gráfico na linha de base do texto.)</span><span class="sxs-lookup"><span data-stu-id="79fca-p142">In the code samples [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML) and [Word-Add-in-Get-Set-EditOpen-XML](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML), the text box and shape are setup using different types of text wrapping and positioning settings. (Also be aware that the image examples in those code samples are setup using in line with text formatting, which positions a graphic object on the text baseline.)</span></span>

<span data-ttu-id="79fca-p143">A forma nesses exemplos de código é posicionada em relação às margens direita e inferior da página. O posicionamento relativo permite fazer a coordenação mais facilmente com a configuração de documento desconhecida do usuário, pois ela se ajustará às margens do usuário e haverá menos risco de causar uma aparência estranha devido às configurações de tamanho do papel, orientação ou margem. Para manter as configurações de posicionamento relativas ao inserir um objeto gráfico, você deve manter a marca de parágrafo (w:p) em que o posicionamento (conhecido no Word como uma âncora) é armazenado. Se inserir o conteúdo em uma marca de parágrafo existente em vez de incluir a sua próprio, você poderá manter a mesma aparência inicial, mas muitos tipos de referências relativas que habilitam o posicionamento a se ajustar automaticamente ao layout do usuário poderão ser perdidos.</span><span class="sxs-lookup"><span data-stu-id="79fca-p143">The shape in those code samples is positioned relative to the right and bottom page margins. Relative positioning lets you more easily coordinate with a user's unknown document setup because it will adjust to the user's margins and run less risk of looking awkward because of paper size, orientation, or margin settings. To retain relative positioning settings when you insert a graphic object, you must retain the paragraph mark (w:p) in which the positioning (known in Word as an anchor) is stored. If you insert the content into an existing paragraph mark rather than including your own, you may be able to retain the same initial visual, but many types of relative references that enable the positioning to automatically adjust to the user's layout may be lost.</span></span>


### <a name="working-with-content-controls"></a><span data-ttu-id="79fca-299">Trabalho com controles de conteúdo</span><span class="sxs-lookup"><span data-stu-id="79fca-299">Working with content controls</span></span>

<span data-ttu-id="79fca-300">Os controles de conteúdo são um recurso importante no Word que pode aprimorar consideravelmente a capacidade do suplemento para o Word de várias maneiras, incluindo permitindo-lhe inserir o conteúdo em locais designados no documento, em vez de apenas na seleção.</span><span class="sxs-lookup"><span data-stu-id="79fca-300">Content controls are an important feature in Word that can greatly enhance the power of your add-in for Word in multiple ways, including giving you the ability to insert content at designated places in the document rather than only at the selection.</span></span>

<span data-ttu-id="79fca-301">No Word, localize os controles de conteúdo na guia Desenvolvedor da faixa de opções, conforme mostrado aqui na Figura 15.</span><span class="sxs-lookup"><span data-stu-id="79fca-301">In Word, find content controls on the Developer tab of the ribbon, as shown here in Figure 15.</span></span>


<span data-ttu-id="79fca-302">*Figura 15. O grupo Controles na guia Desenvolvedor no Word*</span><span class="sxs-lookup"><span data-stu-id="79fca-302">*Figure 15. The Controls group on the Developer tab in Word*</span></span>

![Grupo de Controles de conteúdo na faixa de opções do Word.](../images/office15-app-create-wd-app-using-ooxml-fig15.png)

<span data-ttu-id="79fca-304">Os tipos de controles de conteúdo no Word incluem RTF, texto sem formatação, imagem, galeria de blocos de construção, caixa de seleção, lista suspensa, caixa de combinação, seletor de data e seção de repetição.</span><span class="sxs-lookup"><span data-stu-id="79fca-304">Types of content controls in Word include rich text, plain text, picture, building block gallery, check box, dropdown list, combo box, date picker, and repeating section.</span></span>



- <span data-ttu-id="79fca-305">Use o comando **Propriedades**, mostrado na Figura 15, para editar o título do controle e para definir preferências, como ocultar o contêiner de controle.</span><span class="sxs-lookup"><span data-stu-id="79fca-305">Use the  **Properties** command, shown in Figure 15, to edit the title of the control and to set preferences such as hiding the control container.</span></span>

- <span data-ttu-id="79fca-306">Habilite **Modo de Design** para editar o conteúdo de espaço reservado no controle.</span><span class="sxs-lookup"><span data-stu-id="79fca-306">Enable  **Design Mode** to edit placeholder content in the control.</span></span>

<span data-ttu-id="79fca-p144">Se o suplemento funciona com um modelo do Word, você pode incluir controles no modelo para aprimorar o comportamento do conteúdo. Você também pode usar uma associação de dados XML em um documento do Word para associar controles de conteúdo a dados, como propriedades de documento, para preencher facilmente formulários ou realizar tarefas semelhantes. (Localize os controles que já estão associados a propriedades internas do documento no Word na guia **Inserir** em **Partes Rápidas**.)</span><span class="sxs-lookup"><span data-stu-id="79fca-p144">If your add-in works with a Word template, you can include controls in that template to enhance the behavior of the content. You can also use XML data binding in a Word document to bind content controls to data, such as document properties, for easy form completion or similar tasks. (Find controls that are already bound to built-in document properties in Word on the  **Insert** tab, under **Quick Parts**.)</span></span>

<span data-ttu-id="79fca-p145">Ao usar controles de conteúdo com o suplemento, você também pode expandir muito as opções para o que o suplemento pode fazer usando um tipo diferente de associação. Você pode associar a um controle de conteúdo de dentro do suplemento e, depois, escrever conteúdo para a associação em vez de para a seleção ativa.</span><span class="sxs-lookup"><span data-stu-id="79fca-p145">When you use content controls with your add-in, you can also greatly expand the options for what your add-in can do using a different type of binding. You can bind to a content control from within the add-in and then write content to the binding rather than to the active selection.</span></span>



> [!NOTE]
> <span data-ttu-id="79fca-p146">Não confunda a associação de dados XML no Word com a capacidade de associar a um controle por meio do suplemento. Esses são recursos completamente separados. No entanto, você pode incluir controles de conteúdo nomeados no conteúdo que inserir por meio do suplemento usando a coerção de OOXML e usar código no suplemento para associar a esses controles.</span><span class="sxs-lookup"><span data-stu-id="79fca-p146">Don't confuse XML data binding in Word with the ability to bind to a control via your add-in. These are completely separate features. However, you can include named content controls in the content you insert via your add-in using OOXML coercion and then use code in the add-in to bind to those controls.</span></span>

<span data-ttu-id="79fca-p147">Além disso, lembre-se de que associação de dados XML e o Office.js podem interagir com partes XML personalizadas no aplicativo. Portanto, é possível integrar essas poderosas ferramentas. Para saber mais sobre como trabalhar com partes XML personalizadas na API JavaScript para Office, confira a seção [Recursos adicionais](#see-also) deste tópico.</span><span class="sxs-lookup"><span data-stu-id="79fca-p147">Also be aware that both XML data binding and Office.js can interact with custom XML parts in your app, so it is possible to integrate these powerful tools. To learn about working with custom XML parts in the Office JavaScript API, see the [Additional resources](#see-also) section of this topic.</span></span>

<span data-ttu-id="79fca-p148">O trabalho com associações no suplemento do Word é abordado na próxima seção do tópico. Primeiro, vamos conferir um exemplo do Office Open XML necessário para inserir um controle de conteúdo RTF que você pode associar usando o suplemento.</span><span class="sxs-lookup"><span data-stu-id="79fca-p148">Working with bindings in your Word add-in is covered in the next section of the topic. First, let's take a look at an example of the Office Open XML required for inserting a rich text content control that you can bind to using your add-in.</span></span>



> [!IMPORTANT]
> <span data-ttu-id="79fca-319">Controles RTF são o único tipo de controle de conteúdo que você pode usar para associar a um controle de conteúdo de dentro do suplemento.</span><span class="sxs-lookup"><span data-stu-id="79fca-319">Rich text controls are the only type of content control you can use to bind to a content control from within your add-in.</span></span>




```XML
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" >
        <w:body>
          <w:p/>
          <w:sdt>
              <w:sdtPr>
                <w:alias w:val="MyContentControlTitle"/>
                <w:id w:val="1382295294"/>
                <w15:appearance w15:val="hidden"/>
                <w:showingPlcHdr/>
              </w:sdtPr>
              <w:sdtContent>
                <w:p>
                  <w:r>
                  <w:t>[This text is inside a content control that has its container hidden. You can bind to a content control to add or interact with content at a specified location in the document.]</w:t>
                </w:r>
                </w:p>
              </w:sdtContent>
            </w:sdt>
          </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
 </pkg:package>
```

<span data-ttu-id="79fca-320">Como já mencionado, os controles de conteúdo, como texto formatado, não exigem partes de documento adicionais. Portanto, somente editadas versões das partes .rels e document.xml são incluídas aqui.</span><span class="sxs-lookup"><span data-stu-id="79fca-320">As already mentioned, content controls, like formatted text, don't require additional document parts, so only edited versions of the .rels and document.xml parts are included here.</span></span>

<span data-ttu-id="79fca-p149">A marca **w:sdt** que você vê no corpo de document.xml representa o controle de conteúdo. Se gerar a marcação do Office Open XML para um controle de conteúdo, você verá que vários atributos foram removidos do exemplo, incluindo a marca e as propriedades de parte de documento. Somente elementos essenciais (e alguns de práticas recomendadas) foram mantidos, incluindo o seguinte:</span><span class="sxs-lookup"><span data-stu-id="79fca-p149">The **w:sdt** tag that you see within the document.xml body represents the content control. If you generate the Office Open XML markup for a content control, you'll see that several attributes have been removed from this example, including the tag and document part properties. Only essential (and a couple of best practice) elements have been retained, including the following:</span></span>



- <span data-ttu-id="79fca-p150">O **alias** é a propriedade de título da caixa de diálogo Propriedades de Controle de Conteúdo no Word. Essa é uma propriedade necessária (representando o nome do item) se você planeja associar ao controle de dentro do suplemento.</span><span class="sxs-lookup"><span data-stu-id="79fca-p150">The  **alias** is the title property from the Content Control Properties dialog box in Word. This is a required property (representing the name of the item) if you plan to bind to the control from within your add-in.</span></span>

- <span data-ttu-id="79fca-p151">A **id** exclusiva é uma propriedade necessária. Se você associar ao controle de dentro do suplemento, a ID será a propriedade que a vinculação usa no documento para identificar o controle de conteúdo nomeado aplicável.</span><span class="sxs-lookup"><span data-stu-id="79fca-p151">The unique **id** is a required property. If you bind to the control from within your add-in, the ID is the property the binding uses in the document to identify the applicable named content control.</span></span>

- <span data-ttu-id="79fca-p152">O atributo **aparência** é usado para ocultar o contêiner de controle, para gerar uma aparência mais limpa. Esse é um novo recurso no Word 2013, como você pode ver pelo uso do namespace w15. Como essa propriedade é usada, o namespace w15 é mantida no início da parte document.xml.</span><span class="sxs-lookup"><span data-stu-id="79fca-p152">The  **appearance** attribute is used to hide the control container, for a cleaner look. This feature was introduced in Word 2013, as you see by the use of the w15 namespace. Because this property is used, the w15 namespace is retained at the start of the document.xml part.</span></span>

- <span data-ttu-id="79fca-p153">O atributo **showingPlcHdr** é uma configuração opcional que define o conteúdo padrão que você inclui no controle (texto, neste exemplo) como conteúdo de espaço reservado. Portanto, se o usuário clica ou toca na área de controle, todo o conteúdo é selecionado, em vez de se comportar como conteúdo editável no qual o usuário pode fazer alterações.</span><span class="sxs-lookup"><span data-stu-id="79fca-p153">The  **showingPlcHdr** attribute is an optional setting that sets the default content you include inside the control (text in this example) as placeholder content. So, if the user clicks or taps in the control area, the entire content is selected rather than behaving like editable content in which the user can make changes.</span></span>

- <span data-ttu-id="79fca-p154">Embora a marca de parágrafo vazia (**w:p /**) que precede a marca **sdt** não seja necessária para adicionar um controle de conteúdo (e adicionará espaço vertical acima do controle no documento do Word), ela garante que o controle seja colocado em seu próprio parágrafo. Isso pode ser importante, dependendo do tipo e da formatação do conteúdo será adicionado ao controle.</span><span class="sxs-lookup"><span data-stu-id="79fca-p154">Although the empty paragraph mark ( **w:p/** ) that precedes the **sdt** tag is not required for adding a content control (and will add vertical space above the control in the Word document), it ensures that the control is placed in its own paragraph. This may be important, depending upon the type and formatting of content that will be added in the control.</span></span>

- <span data-ttu-id="79fca-335">Se você pretende associar ao controle, o conteúdo padrão para o controle (o que está dentro da marca **sdtContent**) deve incluir pelo menos um parágrafo completo (como neste exemplo), para que a associação aceite o conteúdo avançado com vários parágrafos.</span><span class="sxs-lookup"><span data-stu-id="79fca-335">If you intend to bind to the control, the default content for the control (what's inside the **sdtContent** tag) must include at least one complete paragraph (as in this example), in order for your binding to accept multi-paragraph rich content.</span></span>



> [!NOTE]
> <span data-ttu-id="79fca-p155">O atributo de parte de documento que foi removido desta marca de exemplo **w:sdt** pode aparecer em um controle de conteúdo para fazer referência a uma parte separada no pacote em que as informações de conteúdo de espaço reservado podem ser armazenadas (partes localizados em um diretório de glossário no pacote do Office Open XML). Embora parte de documento seja o termo usado para partes XML (ou seja, arquivos) dentro de um pacote do Office Open XML, o termo partes de documento, conforme usado na propriedade sdt, refere-se ao mesmo termo no Word que é usado para descrever alguns tipos de conteúdo, incluindo blocos de construção e partes rápidas de propriedade de documento (por exemplo, controles associados a dados XML internos). Se houver partes em um diretório de glossário no pacote do Office Open XML, talvez você precise mantê-las se o conteúdo que estiver inserindo incluir esses recursos. Para um controle de conteúdo típico que você pretende usar para associar do suplemento, elas não são necessárias. Lembre-se apenas de que, se você excluir as partes de glossário do pacote, também deverá remover o atributo de parte de documento da marca w:sdt.</span><span class="sxs-lookup"><span data-stu-id="79fca-p155">The document part attribute that was removed from this sample **w:sdt** tag may appear in a content control to reference a separate part in the package where placeholder content information can be stored (parts located in a glossary directory in the Office Open XML package). Although document part is the term used for XML parts (that is, files) within an Office Open XML package, the term document parts as used in the sdt property refers to the same term in Word that is used to describe some content types including building blocks and document property quick parts (for example, built-in XML data-bound controls). If you see parts under a glossary directory in your Office Open XML package, you may need to retain them if the content you're inserting includes these features. For a typical content control that you intend to use to bind to from your add-in, they're not required. Just remember that, if you do delete the glossary parts from the package, you must also remove the document part attribute from the w:sdt tag.</span></span>

<span data-ttu-id="79fca-341">A próxima seção abordará como criar e usar associações no suplemento do Word.</span><span class="sxs-lookup"><span data-stu-id="79fca-341">The next section will discuss how to create and use bindings in your Word add-in.</span></span>


## <a name="inserting-content-at-a-designated-location"></a><span data-ttu-id="79fca-342">Inserir conteúdo em um local designado</span><span class="sxs-lookup"><span data-stu-id="79fca-342">Inserting content at a designated location</span></span>


<span data-ttu-id="79fca-p156">Já vimos como inserir o conteúdo na seleção ativa em um documento do Word. Se associar a um controle de conteúdo nomeado no documento, você poderá inserir qualquer um dos mesmos tipos de conteúdo no controle.</span><span class="sxs-lookup"><span data-stu-id="79fca-p156">We've already looked at how to insert content at the active selection in a Word document. If you bind to a named content control that's in the document, you can insert any of the same content types into that control.</span></span> 

<span data-ttu-id="79fca-345">Então, quando convém usar essa abordagem?</span><span class="sxs-lookup"><span data-stu-id="79fca-345">So when might you want to use this approach?</span></span>


- <span data-ttu-id="79fca-346">Quando você precisa adicionar ou substituir conteúdo em locais específicos em um modelo, como para preencher partes do documento de um banco de dados</span><span class="sxs-lookup"><span data-stu-id="79fca-346">When you need to add or replace content at specified locations in a template, such as to populate portions of the document from a database</span></span>

- <span data-ttu-id="79fca-347">Quando você quer a opção de substituir o conteúdo que está inserindo na seleção ativa, como para fornecer opções de elemento de design ao usuário</span><span class="sxs-lookup"><span data-stu-id="79fca-347">When you want the option to replace content that you're inserting at the active selection, such as to provide design element options to the user</span></span>

- <span data-ttu-id="79fca-348">Quando você quer que o usuário adicione dados no documento que você possa acessar para uso com o suplemento, como para preencher campos no painel de tarefas com base em informações que o usuário adiciona ao documento</span><span class="sxs-lookup"><span data-stu-id="79fca-348">When you want the user to add data in the document that you can access for use with your add-in, such as to populate fields in the task pane based upon information the user adds in the document</span></span>

<span data-ttu-id="79fca-349">Baixe o código de exemplo [Word-Add-in-JavaScript-AddPopulateBindings](https://github.com/OfficeDev/Word-Add-in-JavaScript-AddPopulateBindings), que fornece um exemplo de como inserir e associar a um controle de conteúdo e como preencher a associação.</span><span class="sxs-lookup"><span data-stu-id="79fca-349">Download the code sample [Word-Add-in-JavaScript-AddPopulateBindings](https://github.com/OfficeDev/Word-Add-in-JavaScript-AddPopulateBindings), which provides a working example of how to insert and bind to a content control, and how to populate the binding.</span></span>


### <a name="add-and-bind-to-a-named-content-control"></a><span data-ttu-id="79fca-350">Adicionar e associar a um controle de conteúdo nomeado</span><span class="sxs-lookup"><span data-stu-id="79fca-350">Add and bind to a named content control</span></span>


<span data-ttu-id="79fca-351">Ao examinar o JavaScript a seguir, considere estes requisitos:</span><span class="sxs-lookup"><span data-stu-id="79fca-351">As you examine the JavaScript that follows, consider these requirements:</span></span>


- <span data-ttu-id="79fca-352">Conforme mencionado anteriormente, você deve usar um controle de conteúdo avançado para associar ao controle do suplemento do Word.</span><span class="sxs-lookup"><span data-stu-id="79fca-352">As previously mentioned, you must use a rich text content control in order to bind to the control from your Word add-in.</span></span>

- <span data-ttu-id="79fca-p157">O controle de conteúdo deve ter um nome (esse é o campo **Título** na caixa de diálogo Propriedades de Controle de Conteúdo, que corresponde à marca **alias** na marcação do Office Open XML). Isso é como o código identifica onde colocar a associação.</span><span class="sxs-lookup"><span data-stu-id="79fca-p157">The content control must have a name (this is the  **Title** field in the Content Control Properties dialog box, which corresponds to the **Alias** tag in the Office Open XML markup). This is how the code identifies where to place the binding.</span></span>

- <span data-ttu-id="79fca-p158">Você pode ter vários controles nomeados e associá-los conforme necessário. Use um nome de controle de conteúdo, uma ID de controle de conteúdo e uma ID de associação exclusivos.</span><span class="sxs-lookup"><span data-stu-id="79fca-p158">You can have several named controls and bind to them as needed. Use a unique content control name, unique content control ID, and a unique binding ID.</span></span>


```js
function addAndBindControl() {
    Office.context.document.bindings.addFromNamedItemAsync("MyContentControlTitle", "text", { id: 'myBinding' }, function (result) {
        if (result.status == "failed") {
            if (result.error.message == "The named item does not exist.")
                var myOOXMLRequest = new XMLHttpRequest();
                var myXML;
                myOOXMLRequest.open('GET', '../../Snippets_BindAndPopulate/ContentControl.xml', false);
                myOOXMLRequest.send();
                if (myOOXMLRequest.status === 200) {
                    myXML = myOOXMLRequest.responseText;
                }
                Office.context.document.setSelectedDataAsync(myXML, { coercionType: 'ooxml' }, function (result) {
                    Office.context.document.bindings.addFromNamedItemAsync("MyContentControlTitle", "text", { id: 'myBinding' });
                });
        }
    });
}
```

<span data-ttu-id="79fca-357">O código mostrado aqui realiza as seguintes etapas:</span><span class="sxs-lookup"><span data-stu-id="79fca-357">The code shown here takes the following steps:</span></span>


- <span data-ttu-id="79fca-358">Tenta associar ao controle de conteúdo nomeado, usando [addFromNamedItemAsync](/javascript/api/office/office.bindings#addfromnameditemasync-itemname--bindingtype--options--callback-).</span><span class="sxs-lookup"><span data-stu-id="79fca-358">Attempts to bind to the named content control, using [addFromNamedItemAsync](/javascript/api/office/office.bindings#addfromnameditemasync-itemname--bindingtype--options--callback-).</span></span>

  <span data-ttu-id="79fca-p159">Execute esta etapa primeiro se houver uma possibilidade para seu suplemento em que o controle nomeado pode já existir no documento quando o código for executado. Por exemplo, faça isto se o suplemento foi inserido em e salvo com um modelo projetado para funcionar com o suplemento, em que o controle foi colocado anteriormente. Você também precisa fazer isto caso necessite associar a um controle que foi colocado anteriormente pelo suplemento.</span><span class="sxs-lookup"><span data-stu-id="79fca-p159">Take this step first if there is a possible scenario for your add-in where the named control could already exist in the document when the code executes. For example, you'll want to do this if the add-in was inserted into and saved with a template that's been designed to work with the add-in, where the control was placed in advance. You also need to do this if you need to bind to a control that was placed earlier by the add-in.</span></span>

- <span data-ttu-id="79fca-p160">O retorno de chamada na primeira chamada ao método **addFromNamedItemAsync** verifica o status do resultado para ver se a associação falhou porque o item nomeado não existe no documento (ou seja, o controle de conteúdo chamado MyContentControlTitle neste exemplo). Nesse caso, o código adiciona o controle no ponto de seleção ativo (usando **setSelectedDataAsync**) e associa a ele.</span><span class="sxs-lookup"><span data-stu-id="79fca-p160">The callback in the first call to the  **addFromNamedItemAsync** method checks the status of the result to see if the binding failed because the named item doesn't exist in the document (that is, the content control named MyContentControlTitle in this example). If so, the code adds the control at the active selection point (using **setSelectedDataAsync** ) and then binds to it.</span></span>


> [!NOTE]
> <span data-ttu-id="79fca-p161">Como mencionado anteriormente e mostrado no código anterior, o nome do controle de conteúdo é usado para determinar onde criar a associação. No entanto, na marcação do Office Open XML, o código adiciona a associação ao documento usando o nome e o atributo de ID do controle de conteúdo.</span><span class="sxs-lookup"><span data-stu-id="79fca-p161">As mentioned earlier and shown in the preceding code, the name of the content control is used to determine where to create the binding. However, in the Office Open XML markup, the code adds the binding to the document using both the name and the ID attribute of the content control.</span></span>

<span data-ttu-id="79fca-p162">Após a execução de código, se examinar a marcação do documento no qual o suplemento criou associações, você verá duas partes para cada associação. Na marcação do controle de conteúdo em que uma associação foi adicionada (em document.xml), você verá o atributo **w15:webExtensionLinked/**.</span><span class="sxs-lookup"><span data-stu-id="79fca-p162">After code execution, if you examine the markup of the document in which your add-in created bindings, you'll see two parts to each binding. In the markup for the content control where a binding was added (in document.xml), you'll see the attribute  **w15:webExtensionLinked/**.</span></span>

<span data-ttu-id="79fca-p163">Na parte do documento chamada webExtensions1.xml, você verá uma lista das associações que criou. Cada uma delas é identificada usando a ID de associação e o atributo de ID do controle aplicável, como o item a seguir, em que o atributo **appref** é a ID de controle de conteúdo: \*\* **we:binding id="myBinding" type="text" appref="1382295294"/**.</span><span class="sxs-lookup"><span data-stu-id="79fca-p163">In the document part named webExtensions1.xml, you'll see a list of the bindings you've created. Each is identified using the binding ID and the ID attribute of the applicable control, such as the following, where the **appref** attribute is the content control ID: \*\* **we:binding id="myBinding" type="text" appref="1382295294"/**.</span></span>


> [!IMPORTANT]
> <span data-ttu-id="79fca-p164">Você deve adicionar a associação no momento em que pretende agir sobre ela. Não inclua a marcação da associação no Office Open XML para inserir o controle de conteúdo, pois o processo de inserção dessa marcação removerá a associação.</span><span class="sxs-lookup"><span data-stu-id="79fca-p164">You must add the binding at the time you intend to act upon it. Don't include the markup for the binding in the Office Open XML for inserting the content control because the process of inserting that markup will strip the binding.</span></span>


### <a name="populate-a-binding"></a><span data-ttu-id="79fca-372">Preencher uma associação</span><span class="sxs-lookup"><span data-stu-id="79fca-372">Populate a binding</span></span>


<span data-ttu-id="79fca-373">O código para gravar conteúdo para uma associação é semelhante ao usado para gravar conteúdo para uma seleção.</span><span class="sxs-lookup"><span data-stu-id="79fca-373">The code for writing content to a binding is similar to that for writing content to a selection.</span></span>


```js
function populateBinding(filename) {
  var myOOXMLRequest = new XMLHttpRequest();
  var myXML;
  myOOXMLRequest.open('GET', filename, false);
  myOOXMLRequest.send();
  if (myOOXMLRequest.status === 200) {
      myXML = myOOXMLRequest.responseText;
  }
  Office.select("bindings#myBinding").setDataAsync(myXML, { coercionType: 'ooxml' });
}
```

<span data-ttu-id="79fca-p165">Assim como ocorre com **setSelectedDataAsync**, você especifica o conteúdo a ser inserido e o tipo de coerção. O único requisito adicional para gravar em uma associação é identificá-la por ID. Observe como a ID de associação usada neste código (bindings#myBinding) corresponde à ID de associação estabelecida (myBinding) quando a associação foi criada na função anterior.</span><span class="sxs-lookup"><span data-stu-id="79fca-p165">As with  **setSelectedDataAsync**, you specify the content to be inserted and the coercion type. The only additional requirement for writing to a binding is to identify the binding by ID. Notice how the binding ID used in this code (bindings#myBinding) corresponds to the binding ID established (myBinding) when the binding was created in the previous function.</span></span>


> [!NOTE]
> <span data-ttu-id="79fca-p166">O código anterior é tudo de que você precisará se estiver preenchendo ou substituindo inicialmente o conteúdo em uma associação. Quando você insere um novo item de conteúdo em um local associado, o conteúdo existente na associação é substituído automaticamente. Confira um exemplo disso no exemplo de código referenciado anteriormente, [Word-Add-in-JavaScript-AddPopulateBindings](https://github.com/OfficeDev/Word-Add-in-JavaScript-AddPopulateBindings), que fornece dois exemplos de conteúdo separados que você pode intercambiar para preencher a mesma associação.</span><span class="sxs-lookup"><span data-stu-id="79fca-p166">The preceding code is all you need whether you are initially populating or replacing the content in a binding. When you insert a new piece of content at a bound location, the existing content in that binding is automatically replaced. Check out an example of this in the previously-referenced code sample [Word-Add-in-JavaScript-AddPopulateBindings](https://github.com/OfficeDev/Word-Add-in-JavaScript-AddPopulateBindings), which provides two separate content samples that you can use interchangeably to populate the same binding.</span></span>


## <a name="adding-objects-that-use-additional-office-open-xml-parts"></a><span data-ttu-id="79fca-380">Adicione objetos que usam partes adicionais do Office Open XML</span><span class="sxs-lookup"><span data-stu-id="79fca-380">Adding objects that use additional Office Open XML parts</span></span>


<span data-ttu-id="79fca-381">Muitos tipos de conteúdo exigem partes adicionais do documento no pacote do Office Open XML, o que significa que fazem referência a informações em outra parte ou o próprio conteúdo é armazenado em uma ou mais partes adicionais e referenciado em document.xml.</span><span class="sxs-lookup"><span data-stu-id="79fca-381">Many types of content require additional document parts in the Office Open XML package, meaning that they either reference information in another part or the content itself is stored in one or more additional parts and referenced in document.xml.</span></span>

<span data-ttu-id="79fca-382">Por exemplo, considere a seguinte situação:</span><span class="sxs-lookup"><span data-stu-id="79fca-382">For example, consider the following:</span></span>


- <span data-ttu-id="79fca-383">O conteúdo que usa estilos para formatação (como o texto com estilo mostrado anteriormente na Figura 2 ou a tabela com estilo mostrada na Figura 9) requer a parte styles.xml.</span><span class="sxs-lookup"><span data-stu-id="79fca-383">Content that uses styles for formatting (such as the styled text shown earlier in Figure 2 or the styled table shown in Figure 9) requires the styles.xml part.</span></span>

- <span data-ttu-id="79fca-384">Imagens (como as mostradas na Figuras 3 e 4) incluem os dados de imagem binários em uma e, às vezes, em duas partes adicionais.</span><span class="sxs-lookup"><span data-stu-id="79fca-384">Images (such as those shown in Figures 3 and 4) include the binary image data in one (and sometimes two) additional parts.</span></span>

- <span data-ttu-id="79fca-385">Diagramas SmartArt (como o que é mostrado na Figura 10) exigem várias partes adicionais para descrever o layout e o conteúdo.</span><span class="sxs-lookup"><span data-stu-id="79fca-385">SmartArt diagrams (such as the one shown in Figure 10) require multiple additional parts to describe the layout and content.</span></span>

- <span data-ttu-id="79fca-386">Gráficos (como o que é mostrado na Figura 11) exigem várias partes adicionais, incluindo sua própria parte de relação (.rels).</span><span class="sxs-lookup"><span data-stu-id="79fca-386">Charts (such as the one shown in Figure 11) require multiple additional parts, including their own relationship (.rels) part.</span></span>

<span data-ttu-id="79fca-p167">Você pode ver exemplos editados da marcação para todos esses tipos de conteúdo no exemplo de código referenciado anteriormente, [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML). Você pode inserir todos esses tipos de conteúdo usando o mesmo código JavaScript mostrado anteriormente (e fornecido nos exemplos de código referenciados) para inserir o conteúdo na seleção ativa e gravar conteúdo em um local específico usando associações.</span><span class="sxs-lookup"><span data-stu-id="79fca-p167">You can see edited examples of the markup for all of these content types in the previously-referenced code sample [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML). You can insert all of these content types using the same JavaScript code shown earlier (and provided in the referenced code samples) for inserting content at the active selection and writing content to a specified location using bindings.</span></span>

<span data-ttu-id="79fca-389">Antes que você explore os exemplos, vamos conferir algumas dicas para trabalhar com cada um desses tipos de conteúdo.</span><span class="sxs-lookup"><span data-stu-id="79fca-389">Before you explore the samples, let's take a look at few tips for working with each of these content types.</span></span>


> [!IMPORTANT]
> <span data-ttu-id="79fca-390">Lembre-se: se mantiver partes adicionais referenciadas em document.xml, você precisará manter document.xml.rels e as definições de relação das partes aplicáveis que está mantendo, como styles.xml ou um arquivo de imagem.</span><span class="sxs-lookup"><span data-stu-id="79fca-390">Remember, if you are retaining any additional parts referenced in document.xml, you will need to retain document.xml.rels and the relationship definitions for the applicable parts you're keeping, such as styles.xml or an image file.</span></span>


### <a name="working-with-styles"></a><span data-ttu-id="79fca-391">Como trabalhar com estilos</span><span class="sxs-lookup"><span data-stu-id="79fca-391">Working with styles</span></span>

<span data-ttu-id="79fca-p168">A mesma abordagem para edição de marcação que vimos no exemplo anterior com texto formatado diretamente é aplicada ao se usar estilos de parágrafo ou estilos de tabela para formatar o conteúdo. No entanto, a marcação para trabalhar com estilos de parágrafo é consideravelmente mais simples. Portanto, esse é o exemplo descrito aqui.</span><span class="sxs-lookup"><span data-stu-id="79fca-p168">The same approach to editing the markup that we looked at for the preceding example with directly-formatted text applies when using paragraph styles or table styles to format your content. However, the markup for working with paragraph styles is considerably simpler, so that is the example described here.</span></span>


#### <a name="editing-the-markup-for-content-using-paragraph-styles"></a><span data-ttu-id="79fca-394">Editar a marcação de conteúdo usando estilos de parágrafo</span><span class="sxs-lookup"><span data-stu-id="79fca-394">Editing the markup for content using paragraph styles</span></span>

<span data-ttu-id="79fca-395">A marcação a seguir representa o conteúdo do corpo para o exemplo de texto com estilo mostrado na Figura 2.</span><span class="sxs-lookup"><span data-stu-id="79fca-395">The following markup represents the body content for the styled text example shown in Figure 2.</span></span>


```XML
<w:body>
  <w:p>
    <w:pPr>
      <w:pStyle w:val="Heading1"/>
    </w:pPr>
    <w:r>
      <w:t>This text is formatted using the Heading 1 paragraph style.</w:t>
    </w:r>
  </w:p>
</w:body>
```


> [!NOTE]
> <span data-ttu-id="79fca-p169">Como você pode ver, a marcação de texto formatado em document.xml é consideravelmente mais simples quando você usa um estilo, pois o estilo contém toda a formatação de parágrafo e fonte que, caso contrário, você precisa referenciar individualmente. No entanto, conforme explicado anteriormente, talvez você queira usar estilos ou formatação direta para fins diferentes: usar formatação direta para especificar a aparência do texto independentemente da formatação no documento do usuário; usar um estilo de parágrafo (particularmente um nome de estilo de parágrafo interno, como Título 1, mostrado aqui) para que a formatação do texto seja automaticamente coordenada com o documento do usuário.</span><span class="sxs-lookup"><span data-stu-id="79fca-p169">As you see, the markup for formatted text in document.xml is considerably simpler when you use a style, because the style contains all of the paragraph and font formatting that you otherwise need to reference individually. However, as explained earlier, you might want to use styles or direct formatting for different purposes: use direct formatting to specify the appearance of your text regardless of the formatting in the user's document; use a paragraph style (particularly a built-in paragraph style name, such as Heading 1 shown here) to have the text formatting automatically coordinate with the user's document.</span></span>

<span data-ttu-id="79fca-p170">O uso de um estilo é um bom exemplo da importância de ler e entender a marcação para o conteúdo que você está inserindo, pois não é explícito que outra parte do documento é referenciada aqui. Se você incluir a definição de estilo na marcação e não incluir a parte styles.xml, as informações de estilo em document.xml serão ignoradas independentemente de esse estilo estar ou não em uso no documento do usuário.</span><span class="sxs-lookup"><span data-stu-id="79fca-p170">Use of a style is a good example of how important it is to read and understand the markup for the content you're inserting, because it's not explicit that another document part is referenced here. If you include the style definition in this markup and don't include the styles.xml part, the style information in document.xml will be ignored regardless of whether or not that style is in use in the user's document.</span></span>

<span data-ttu-id="79fca-400">No entanto, se analisar a parte styles.xml, verá que apenas uma pequena parte dessa longa marcação é necessária ao editar a marcação para uso no suplemento:</span><span class="sxs-lookup"><span data-stu-id="79fca-400">However, if you take a look at the styles.xml part, you'll see that only a small portion of this long piece of markup is required when editing markup for use in your add-in:</span></span>


- <span data-ttu-id="79fca-p171">A parte styles.xml inclui vários namespaces por padrão. Se você estiver mantendo apenas as informações de estilo necessárias para o conteúdo, na maioria dos casos, só precisará manter o namespace **xmlns:w**.</span><span class="sxs-lookup"><span data-stu-id="79fca-p171">The styles.xml part includes several namespaces by default. If you are only retaining the required style information for your content, in most cases you only need to keep the **xmlns:w** namespace.</span></span>

- <span data-ttu-id="79fca-403">O conteúdo da marca **w:docDefaults** que fica no início da parte styles será ignorado quando a marcação for inserida por meio do suplemento e pode ser removido.</span><span class="sxs-lookup"><span data-stu-id="79fca-403">The **w:docDefaults** tag content that falls at the top of the styles part will be ignored when your markup is inserted via the add-in and can be removed.</span></span>

- <span data-ttu-id="79fca-p172">A maior marcação em uma parte styles.xml é para a marca **w:latentStyles** que aparece depois de docDefaults, que fornece informações (como atributos de aparência para o painel Estilos e a galeria de Estilos) para todos os estilos disponíveis. Essas informações também serão ignoradas ao se inserir conteúdo por meio do suplemento e, assim, podem ser removidas.</span><span class="sxs-lookup"><span data-stu-id="79fca-p172">The largest piece of markup in a styles.xml part is for the **w:latentStyles** tag that appears after docDefaults, which provides information (such as appearance attributes for the Styles pane and Styles gallery) for every available style. This information is also ignored when inserting content via your add-in and so it can be removed.</span></span>

- <span data-ttu-id="79fca-p173">Após as informações de estilos latentes, você vê uma definição de cada estilo em uso no documento a partir do qual a marcação foi gerada. Isso inclui alguns estilos padrão que estão em uso quando você cria um novo documento e podem não ser relevantes ao conteúdo. Você pode excluir as definições de estilos que não são usadas pelo conteúdo.</span><span class="sxs-lookup"><span data-stu-id="79fca-p173">Following the latent styles information, you see a definition for each style in use in the document from which you're markup was generated. This includes some default styles that are in use when you create a new document and may not be relevant to your content. You can delete the definitions for any styles that aren't used by your content.</span></span>


   > [!NOTE]
   > <span data-ttu-id="79fca-p174">Cada estilo de título interno tem um estilo Char associado que é uma versão de estilo de caractere do mesmo formato do título. A menos que tenha aplicado o estilo de título como um estilo de caractere, você pode removê-lo. Se o estilo for usado como um estilo de caractere, ele aparecerá em document.xml em uma marca de propriedades de execução (**w:rPr**) em vez de uma marca de propriedades de parágrafo (**w:pPr**). Isso só deverá ocorrer se você tiver aplicado o estilo apenas a parte de um parágrafo, mas poderá ocorrer inadvertidamente se o estilo tiver sido aplicado de forma incorreta.</span><span class="sxs-lookup"><span data-stu-id="79fca-p174">Each built-in heading style has an associated Char style that is a character style version of the same heading format. Unless you've applied the heading style as a character style, you can remove it. If the style is used as a character style, it appears in document.xml in a run properties tag ( **w:rPr** ) rather than a paragraph properties ( **w:pPr** ) tag. This should only be the case if you've applied the style to just part of a paragraph, but it can occur inadvertently if the style was incorrectly applied.</span></span>


- <span data-ttu-id="79fca-p175">Se estiver usando um estilo interno para o conteúdo, você não precisará incluir uma definição completa. Você só deve incluir o nome do estilo, a ID do estilo e pelo menos um atributo de formatação para que o Office Open XML com coerção aplique o estilo ao conteúdo durante a inserção.</span><span class="sxs-lookup"><span data-stu-id="79fca-p175">If you're using a built-in style for your content, you don't have to include a full definition. You only must include the style name, style ID, and at least one formatting attribute in order for the coerced Office Open XML to apply the style to your content upon insertion.</span></span>

    <span data-ttu-id="79fca-p176">No entanto, a prática recomendada é incluir uma definição de estilo completa (mesmo que seja o padrão para os estilos internos). Se um estilo já estiver sendo usado no documento de destino, seu conteúdo adotará a definição do residente para o estilo, independentemente de você incluir no styles.xml. Se o estilo ainda não estiver sendo usado no documento de destino, seu conteúdo usará a definição de estilo que você forneceu na marcação.</span><span class="sxs-lookup"><span data-stu-id="79fca-p176">However, it's a best practice to include a complete style definition (even if it's the default for built-in styles). If a style is already in use in the destination document, your content will take on the resident definition for the style, regardless of what you include in styles.xml. If the style isn't yet in use in the destination document, your content will use the style definition you provide in the markup.</span></span>

<span data-ttu-id="79fca-418">Portanto, por exemplo, o único conteúdo que é preciso manter da parte styles.xml para o texto de exemplo mostrado na Figura 2, que é formatado com o estilo Título 1, é indicado a seguir.</span><span class="sxs-lookup"><span data-stu-id="79fca-418">So, for example, the only content we needed to retain from the styles.xml part for the sample text shown in Figure 2, which is formatted using Heading 1 style, is the following.</span></span>


> [!NOTE]
> <span data-ttu-id="79fca-419">Uma definição completa do Word para o estilo Título 1 foi mantida neste exemplo.</span><span class="sxs-lookup"><span data-stu-id="79fca-419">A complete Word definition for the Heading 1 style has been retained in this example.</span></span>




```XML
<pkg:part pkg:name="/word/styles.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml">
  <pkg:xmlData>
    <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" >
      <w:style w:type="paragraph" w:styleId="Heading1">
        <w:name w:val="heading 1"/>
        <w:basedOn w:val="Normal"/>
        <w:next w:val="Normal"/>
        <w:link w:val="Heading1Char"/>
        <w:uiPriority w:val="9"/>
        <w:qFormat/>
        <w:pPr>
          <w:keepNext/>
          <w:keepLines/>
          <w:spacing w:before="240" w:after="0" w:line="259" w:lineRule="auto"/>
          <w:outlineLvl w:val="0"/>
        </w:pPr>
        <w:rPr>
          <w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>
          <w:color w:val="2E74B5" w:themeColor="accent1" w:themeShade="BF"/>
          <w:sz w:val="32"/>
          <w:szCs w:val="32"/>
        </w:rPr>
      </w:style>
    </w:styles>
  </pkg:xmlData>
</pkg:part>
```


#### <a name="editing-the-markup-for-content-using-table-styles"></a><span data-ttu-id="79fca-420">Como editar a marcação de conteúdo usando estilos de tabela</span><span class="sxs-lookup"><span data-stu-id="79fca-420">Editing the markup for content using table styles</span></span>


<span data-ttu-id="79fca-p177">Quando o conteúdo usa um estilo de tabela, você precisa da mesma parte relativa de styles.xml conforme descrito para trabalhar com estilos de parágrafo. Ou seja, você só precisa manter as informações do estilo que está usando no conteúdo e deve incluir o nome, a ID e pelo menos um atributo de formatação, mas é melhor incluir uma definição de estilo completa para lidar com todos os cenários de usuário possíveis.</span><span class="sxs-lookup"><span data-stu-id="79fca-p177">When your content uses a table style, you need the same relative part of styles.xml as described for working with paragraph styles. That is, you only need to retain the information for the style you're using in your content, and you must include the name, ID, and at least one formatting attribute, but are better off including a complete style definition to address all potential user scenarios.</span></span>

<span data-ttu-id="79fca-423">No entanto, ao conferir a marcação da tabela em document.xml e da definição de estilo de tabela em styles.xml, você vê muito mais marcação do que ao trabalhar com estilos de parágrafo.</span><span class="sxs-lookup"><span data-stu-id="79fca-423">However, when you look at the markup both for your table in document.xml and for your table style definition in styles.xml, you see enormously more markup than when working with paragraph styles.</span></span>


- <span data-ttu-id="79fca-p178">Em document.xml, a formatação é aplicada por célula, mesmo que esteja incluída em um estilo. O uso de um estilo de tabela não reduz o volume de marcação. A vantagem de usar estilos de tabela para o conteúdo é facilitar a atualização e coordenar facilmente a aparência de várias tabelas.</span><span class="sxs-lookup"><span data-stu-id="79fca-p178">In document.xml, formatting is applied by cell even if it's included in a style. Using a table style won't reduce the volume of markup. The benefit of using table styles for the content is for easy updating and easily coordinating the look of multiple tables.</span></span>

- <span data-ttu-id="79fca-427">Em styles.xml, você verá uma grande quantidade de marcação para um único estilo de tabela, pois estilos de tabela incluem vários tipos de atributos de formatação possíveis para cada uma das várias áreas da tabela, como a tabela inteira, linhas de título, linhas e colunas em faixas pares e ímpares (separadamente), a primeira coluna etc.</span><span class="sxs-lookup"><span data-stu-id="79fca-427">In styles.xml, you'll see a substantial amount of markup for a single table style as well, because table styles include several types of possible formatting attributes for each of several table areas, such as the entire table, heading rows, odd and even banded rows and columns (separately), the first column, etc.</span></span>


### <a name="working-with-images"></a><span data-ttu-id="79fca-428">Trabalhar com imagens</span><span class="sxs-lookup"><span data-stu-id="79fca-428">Working with images</span></span>


<span data-ttu-id="79fca-p179">A marcação de uma imagem inclui uma referência a pelo menos uma parte que inclua os dados binários para descrever a imagem. Para uma imagem complexa, isso pode consistir em centenas de páginas de marcação, e você não pode editá-la. Como nunca precisa alterar a(s) parte(s) binária(s), você poderá simplesmente recolhê-la(s) se estiver usando um editor estruturado como o Visual Studio, para que ainda possa examinar e editar facilmente o restante do pacote.</span><span class="sxs-lookup"><span data-stu-id="79fca-p179">The markup for an image includes a reference to at least one part that includes the binary data to describe your image. For a complex image, this can be hundreds of pages of markup and you can't edit it. Since you don't ever have to touch the binary part(s), you can simply collapse it if you're using a structured editor such as Visual Studio, so that you can still easily review and edit the rest of the package.</span></span>

<span data-ttu-id="79fca-p180">Se examinar a marcação de exemplo da imagem simples mostrada anteriormente na Figura 3, disponível no exemplo de código mencionado anteriormente, [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML), você verá que a marcação da imagem em document.xml inclui informações de tamanho e posição, bem como uma referência de relação para a parte que contém os dados de imagem binários. Essa referência é incluída na marca **a:blip**, da seguinte maneira:</span><span class="sxs-lookup"><span data-stu-id="79fca-p180">If you check out the example markup for the simple image shown earlier in Figure 3, available in the previously-referenced code sample [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML), you'll see that the markup for the image in document.xml includes size and position information as well as a relationship reference to the part that contains the binary image data. That reference is included in the **a:blip** tag, as follows:</span></span>




```XML
<a:blip r:embed="rId4" cstate="print">
```

<span data-ttu-id="79fca-p181">Lembre-se de que, como uma referência de relação é usada explicitamente (**r:embed="rID4"**) e essa parte relacionada é necessária para processar a imagem, se você não incluir os dados binários no pacote do Office Open XML, receberá um erro. Isso é diferente do styles.xml, explicado anteriormente, que não gerará um erro se for omitido, pois a relação não é referenciada explicitamente e a relação é para uma parte que fornece atributos ao conteúdo (formatação), em vez de fazer parte do próprio conteúdo.</span><span class="sxs-lookup"><span data-stu-id="79fca-p181">Be aware that, because a relationship reference is explicitly used ( **r:embed="rID4"** ) and that related part is required in order to render the image, if you don't include the binary data in your Office Open XML package, you will get an error. This is different from styles.xml, explained previously, which won't throw an error if omitted since the relationship is not explicitly referenced and the relationship is to a part that provides attributes to the content (formatting) rather than being part of the content itself.</span></span>


> [!NOTE]
> <span data-ttu-id="79fca-p182">Ao examinar a marcação, observe os namespaces adicionais usados na marca a:blip. Você verá em document.xml que o namespace **xlmns:a** (o namespace drawingML principal) é colocado dinamicamente no início do uso de referências de drawingML, em vez de no início da parte document.xml. No entanto, o namespace de relações (r) deve ser mantido onde aparece no início de document.xml. Verifique se a marcação de imagem tem requisitos de namespace adicionais. Lembre-se de que você não precisa memorizar quais tipos de conteúdo exigem quais namespaces. Você pode distingui-los facilmente examinando os prefixos das marcas em todo o document.xml.</span><span class="sxs-lookup"><span data-stu-id="79fca-p182">When you review the markup, notice the additional namespaces used in the a:blip tag. You'll see in document.xml that the  **xlmns:a** namespace (the main drawingML namespace) is dynamically placed at the beginning of the use of drawingML references rather than at the top of the document.xml part. However, the relationships namespace (r) must be retained where it appears at the start of document.xml. Check your picture markup for additional namespace requirements. Remember that you don't have to memorize which types of content require what namespaces, you can easily tell by reviewing the prefixes of the tags throughout document.xml.</span></span>


### <a name="understanding-additional-image-parts-and-formatting"></a><span data-ttu-id="79fca-441">Noções básicas sobre partes de imagem e formatação adicionais</span><span class="sxs-lookup"><span data-stu-id="79fca-441">Understanding additional image parts and formatting</span></span>


<span data-ttu-id="79fca-p183">Quando você usa alguns efeitos de formatação do Office na imagem, como para a imagem mostrada na Figura 4, que usa configurações ajustadas de brilho e contraste (além de estilo de imagem), pode ser necessária uma segunda parte de dados binários de uma cópia de formato HD dos dados da imagem. Esse formato HD adicional é necessário para a formatação que é considerada um efeito de camadas, e a referência a ele aparece em document.xml, de forma semelhante ao seguinte:</span><span class="sxs-lookup"><span data-stu-id="79fca-p183">When you use some Office picture formatting effects on your image, such as for the image shown in Figure 4, which uses adjusted brightness and contrast settings (in addition to picture styling), a second binary data part for an HD format copy of the image data may be required. This additional HD format is required for formatting considered a layering effect, and the reference to it appears in document.xml similar to the following:</span></span>


```XML
<a14:imgLayer r:embed="rId5">
```

<span data-ttu-id="79fca-444">Veja a marcação necessária para a imagem formatada mostrada na Figura 4 (que usa efeitos das camadas, entre outras) no exemplo de código [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML).</span><span class="sxs-lookup"><span data-stu-id="79fca-444">See the required markup for the formatted image shown in Figure 4 (which uses layering effects among others) in the [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML) code sample.</span></span>


### <a name="working-with-smartart-diagrams"></a><span data-ttu-id="79fca-445">Trabalhar com diagramas SmartArt</span><span class="sxs-lookup"><span data-stu-id="79fca-445">Working with SmartArt diagrams</span></span>


<span data-ttu-id="79fca-p184">Um diagrama SmartArt tem quatro partes associadas, mas apenas duas são sempre necessárias. Você pode examinar um exemplo de marcação SmartArt no exemplo de código [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML). Primeiro, confira uma breve descrição de cada uma das partes e por que elas são necessárias ou não:</span><span class="sxs-lookup"><span data-stu-id="79fca-p184">A SmartArt diagram has four associated parts, but only two are always required. You can examine an example of SmartArt markup in the [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML) code sample. First, take a look at a brief description of each of the parts and why they are or are not required:</span></span>


> [!NOTE]
> <span data-ttu-id="79fca-449">Se o conteúdo incluir mais de um diagrama, eles serão numerados consecutivamente, substituindo o 1 nos nomes de arquivo listados aqui.</span><span class="sxs-lookup"><span data-stu-id="79fca-449">If your content includes more than one diagram, they will be numbered consecutively, replacing the 1 in the file names listed here.</span></span>


- <span data-ttu-id="79fca-p185">layout1.xml: essa parte é necessária. Ela inclui a definição de marcação para a funcionalidade e aparência de layout.</span><span class="sxs-lookup"><span data-stu-id="79fca-p185">layout1.xml: This part is required. It includes the markup definition for the layout appearance and functionality.</span></span>

- <span data-ttu-id="79fca-p186">data1.xml: essa parte é necessária. Ela inclui os dados em uso na instância do diagrama.</span><span class="sxs-lookup"><span data-stu-id="79fca-p186">data1.xml: This part is required. It includes the data in use in your instance of the diagram.</span></span>

- <span data-ttu-id="79fca-454">drawing1.xml: essa parte nem sempre é necessária, mas, se você aplicar formatação personalizada a elementos na instância de um diagrama, como formatar diretamente formas individuais, talvez seja necessário mantê-la.</span><span class="sxs-lookup"><span data-stu-id="79fca-454">drawing1.xml: This part is not always required but if you apply custom formatting to elements in your instance of a diagram, such as directly formatting individual shapes, you might need to retain it.</span></span>

- <span data-ttu-id="79fca-p187">colors1.xml: essa parte não é necessária. Ela inclui informações de estilo de cor, mas as cores do diagrama serão coordenadas por padrão com as cores do tema de formatação ativo no documento de destino, com base no estilo de cor SmartArt que você aplicar da guia de design Ferramentas SmartArt no Word antes de salvar a marcação do Office Open XML.</span><span class="sxs-lookup"><span data-stu-id="79fca-p187">colors1.xml: This part is not required. It includes color style information, but the colors of your diagram will coordinate by default with the colors of the active formatting theme in the destination document, based on the SmartArt color style you apply from the SmartArt Tools design tab in Word before saving out your Office Open XML markup.</span></span>

- <span data-ttu-id="79fca-p188">quickStyles1.xml: essa parte não é necessária. De forma semelhante à parte de cores, você pode removê-la, pois o diagrama adotará a definição do estilo SmartArt aplicado que está disponível no documento de destino (ou seja, ele será coordenado automaticamente com o tema de formatação no documento de destino).</span><span class="sxs-lookup"><span data-stu-id="79fca-p188">quickStyles1.xml: This part is not required. Similar to the colors part, you can remove this as your diagram will take on the definition of the applied SmartArt style that's available in the destination document (that is, it will automatically coordinate with the formatting theme in the destination document).</span></span>


> [!TIP]
> <span data-ttu-id="79fca-p189">O arquivo SmartArt layout1.xml é um bom exemplo de locais em que talvez você consiga cortar ainda mais a marcação, mas talvez não valha a pena o tempo extra para fazer isso (porque é removida uma pequena quantidade de marcação em relação ao pacote inteiro). Se quiser remover todas as linhas de marcação que puder, você poderá excluir a marca **dgm:sampData** e seu conteúdo. Esses dados de exemplo definem como a visualização de miniatura do diagrama será exibida nas galerias de estilos SmartArt. No entanto, se forem omitidos, dados de exemplo padrão serão usados.</span><span class="sxs-lookup"><span data-stu-id="79fca-p189">The SmartArt layout1.xml file is a good example of places you may be able to further trim your markup but might not be worth the extra time to do so (because it removes such a small amount of markup relative to the entire package). If you would like to get rid of every last line you can of markup, you can delete the **dgm:sampData** tag and its contents. This sample data defines how the thumbnail preview for the diagram will appear in the SmartArt styles galleries. However, if it's omitted, default sample data is used.</span></span>

<span data-ttu-id="79fca-p190">Lembre-se de que a marcação de um diagrama SmartArt em document.xml contém referências de ID de relação para partes de layout, dados, cores e estilos rápidos. Você pode excluir as referências em document.xml das partes de cores e estilos ao excluir essas partes e suas definições de relação (e certamente é uma prática recomendada fazer isso, pois você está excluindo essas relações), mas não receberá um erro se as mantiver, pois não são necessárias para que o diagrama seja inserido em um documento. Localize essas referências em document.xml na marca **dgm:relIds**. Independentemente de você executar esta etapa ou não, mantenha as referências de ID de relação para as partes de dados e layout necessárias.</span><span class="sxs-lookup"><span data-stu-id="79fca-p190">Be aware that the markup for a SmartArt diagram in document.xml contains relationship ID references to the layout, data, colors, and quick styles parts. You can delete the references in document.xml to the colors and styles parts when you delete those parts and their relationship definitions (and it's certainly a best practice to do so, since you're deleting those relationships), but you won't get an error if you leave them, since they aren't required for your diagram to be inserted into a document. Find these references in document.xml in the  **dgm:relIds** tag. Regardless of whether or not you take this step, retain the relationship ID references for the required layout and data parts.</span></span>


### <a name="working-with-charts"></a><span data-ttu-id="79fca-467">Trabalhar com gráficos</span><span class="sxs-lookup"><span data-stu-id="79fca-467">Working with charts</span></span>


<span data-ttu-id="79fca-p191">De forma semelhante aos diagramas SmartArt, os gráficos contêm várias partes adicionais. No entanto, a configuração para os gráficos é um pouco diferente do SmartArt, pois um gráfico tem seu próprio arquivo de relação. A seguir há uma descrição das partes de documento necessárias e removíveis para um gráfico:</span><span class="sxs-lookup"><span data-stu-id="79fca-p191">Similar to SmartArt diagrams, charts contain several additional parts. However, the setup for charts is a bit different from SmartArt, in that a chart has its own relationship file. Following is a description of required and removable document parts for a chart:</span></span>


> [!NOTE]
> <span data-ttu-id="79fca-471">Assim como ocorre com diagramas SmartArt, se o conteúdo incluir mais de um gráfico, eles serão numerados consecutivamente, substituindo o 1 nos nomes de arquivos listados aqui.</span><span class="sxs-lookup"><span data-stu-id="79fca-471">As with SmartArt diagrams, if your content includes more than one chart, they will be numbered consecutively, replacing the 1 in the file names listed here.</span></span>


- <span data-ttu-id="79fca-472">Em document.xml.rels, você verá uma referência à parte necessária que contém os dados que descrevem o gráfico (chart1.xml).</span><span class="sxs-lookup"><span data-stu-id="79fca-472">In document.xml.rels, you'll see a reference to the required part that contains the data that describes the chart (chart1.xml).</span></span>

- <span data-ttu-id="79fca-473">Você também verá um arquivo de relação separada para cada gráfico no pacote do Office Open XML, como chart1.xml.rels.</span><span class="sxs-lookup"><span data-stu-id="79fca-473">You also see a separate relationship file for each chart in your Office Open XML package, such as chart1.xml.rels.</span></span>

    <span data-ttu-id="79fca-p192">Há três arquivos referenciados em chart1.xml.rels, mas apenas um é obrigatório. Eles são os dados binários da pasta de trabalho do Excel (obrigatório) e as cores e partes do estilo (colors1.xml e styles1.xml), que você pode remover.</span><span class="sxs-lookup"><span data-stu-id="79fca-p192">There are three files referenced in chart1.xml.rels, but only one is required. These include the binary Excel workbook data (required) and the color and style parts (colors1.xml and styles1.xml) that you can remove.</span></span>

<span data-ttu-id="79fca-p193">Os gráficos que você pode criar e editar de forma nativa no Word são gráficos do Excel, e seus dados são mantidos em uma planilha do Excel que é inserida como dados binários no pacote do Office Open XML. Assim como as partes de dados binários para imagens, esses dados binários do Excel são necessários, mas não há nada para editar nessa parte. Portanto, você pode simplesmente recolher a parte no editor para evitar ter que rolar manualmente por ela para examinar o restante do pacote do Office Open XML.</span><span class="sxs-lookup"><span data-stu-id="79fca-p193">Charts that you can create and edit natively in Word are Excel charts, and their data is maintained on an Excel worksheet that's embedded as binary data in your Office Open XML package. Like the binary data parts for images, this Excel binary data is required, but there's nothing to edit in this part. So you can just collapse the part in the editor to avoid having to manually scroll through it all to examine the rest of your Office Open XML package.</span></span>

<span data-ttu-id="79fca-p194">No entanto, de forma semelhante ao SmartArt, você pode excluir as partes de cores e estilos. Se você tiver usado os estilos de gráfico e de cor disponíveis para formatar o gráfico, o gráfico adotará a formatação aplicável automaticamente quando for inserido no documento de destino.</span><span class="sxs-lookup"><span data-stu-id="79fca-p194">However, similar to SmartArt, you can delete the colors and styles parts. If you've used the chart styles and color styles available in to format your chart, the chart will take on the applicable formatting automatically when it is inserted into the destination document.</span></span>

<span data-ttu-id="79fca-481">Confira a marcação editada do gráfico de exemplo mostrado na Figura 11 no exemplo de código [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML).</span><span class="sxs-lookup"><span data-stu-id="79fca-481">See the edited markup for the example chart shown in Figure 11 in the [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML) code sample.</span></span>


## <a name="editing-the-office-open-xml-for-use-in-your-task-pane-add-in"></a><span data-ttu-id="79fca-482">Edição do Office Open XML para uso no suplemento de painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="79fca-482">Editing the Office Open XML for use in your task pane add-in</span></span>


<span data-ttu-id="79fca-p195">Você já viu como identificar e editar o conteúdo na marcação. Se a tarefa ainda parecer difícil quando você examinar o enorme pacote do Office Open XML gerado para o documento, veja a seguir um resumo rápido das etapas recomendadas para ajudá-lo a editar o pacote rapidamente:</span><span class="sxs-lookup"><span data-stu-id="79fca-p195">You've already seen how to identify and edit the content in your markup. If the task still seems difficult when you take a look at the massive Office Open XML package generated for your document, following is a quick summary of recommended steps to help you edit that package down quickly:</span></span>


> [!NOTE]
> <span data-ttu-id="79fca-485">Lembre-se de que você pode usar todas as partes .rels no pacote como um mapa para verificar rapidamente se há partes do documento que pode remover.</span><span class="sxs-lookup"><span data-stu-id="79fca-485">Remember that you can use all .rels parts in the package as a map to quickly check for document parts that you can remove.</span></span>


1. <span data-ttu-id="79fca-p196">Abra o arquivo XML compactado no Visual Studio e pressione Ctrl+K, Ctrl+D para formatar o arquivo. Em seguida, use os botões de recolher/expandir à esquerda para recolher as partes que você sabe que precisa remover. Também convém recolher partes longas de que você precisa, mas que sabe que não precisará editar (como os dados binários em base64 para um arquivo de imagem), tornando a verificação visual da marcação mais rápida e fácil.</span><span class="sxs-lookup"><span data-stu-id="79fca-p196">Open the flattened XML file in Visual Studio and press Ctrl+K, Ctrl+D to format the file. Then use the collapse/expand buttons on the left to collapse the parts you know you need to remove. You might also want to collapse long parts you need, but know you won't need to edit (such as the base64 binary data for an image file), making the markup faster and easier to visually scan.</span></span>

2. <span data-ttu-id="79fca-489">Há várias partes do pacote de documento que você quase sempre pode remover ao preparar a marcação do Office Open XML para uso no suplemento.</span><span class="sxs-lookup"><span data-stu-id="79fca-489">There are several parts of the document package that you can almost always remove when you are preparing Office Open XML markup for use in your add-in.</span></span> <span data-ttu-id="79fca-490">Convém começar removendo-as (bem como suas definições de relação associadas), o que reduzirá bastante o pacote de imediato.</span><span class="sxs-lookup"><span data-stu-id="79fca-490">You might want to start by removing these (and their associated relationship definitions), which will greatly reduce the package right away.</span></span> <span data-ttu-id="79fca-491">Elas incluem theme1, fontTable, settings, webSettings, thumbnail, os arquivos de propriedades principal e do suplemento e quaisquer partes de `taskpane` ou de `webExtension`.</span><span class="sxs-lookup"><span data-stu-id="79fca-491">These include the theme1, fontTable, settings, webSettings, thumbnail, both the core and add-in properties files, and any `taskpane` or `webExtension` parts.</span></span>

3. <span data-ttu-id="79fca-p198">Remova as partes que não estão relacionadas ao conteúdo, como notas de rodapé, cabeçalhos ou rodapés dos quais não precisa. Novamente, lembre-se de também excluir suas relações associadas.</span><span class="sxs-lookup"><span data-stu-id="79fca-p198">Remove any parts that don't relate to your content, such as footnotes, headers, or footers that you don't require. Again, remember to also delete their associated relationships.</span></span>

4. <span data-ttu-id="79fca-p199">Examine a parte document.xml.rels para ver se arquivos referenciados nessa parte são necessários para o conteúdo, como um arquivo de imagem, a parte styles ou partes de diagramas SmartArt. Exclua as relações das partes que o conteúdo não requer e confirme se também excluiu a parte associada. Se o conteúdo não exigir nenhuma das partes de documento referenciadas em document.xml.rels, você poderá excluir esse arquivo também.</span><span class="sxs-lookup"><span data-stu-id="79fca-p199">Review the document.xml.rels part to see if any files referenced in that part are required for your content, such as an image file, the styles part, or SmartArt diagram parts. Delete the relationships for any parts your content doesn't require and confirm that you have also deleted the associated part. If your content doesn't require any of the document parts referenced in document.xml.rels, you can delete that file also.</span></span>

5. <span data-ttu-id="79fca-497">Se o conteúdo tiver uma parte .rels adicional (como chart#.xml.rels), examine-o para ver se há outras partes referenciadas que você pode remover (como estilos rápidos para gráficos) e exclua a relação desse arquivo e a parte associada.</span><span class="sxs-lookup"><span data-stu-id="79fca-497">If your content has an additional .rels part (such as chart#.xml.rels), review it to see if there are other parts referenced there that you can remove (such as quick styles for charts) and delete both the relationship from that file as well as the associated part.</span></span>

6. <span data-ttu-id="79fca-p200">Edite document.xml para remover namespaces não referenciados na parte, propriedades da seção se o conteúdo não incluir uma quebra de seção e qualquer marcação que não esteja relacionada ao conteúdo que você deseja inserir. Se você está inserindo formas ou caixas de texto, também convém remover a marcação de fallback ampla.</span><span class="sxs-lookup"><span data-stu-id="79fca-p200">Edit document.xml to remove namespaces not referenced in the part, section properties if your content doesn't include a section break, and any markup that's not related to the content that you want to insert. If inserting shapes or text boxes, you might also want to remove extensive fallback markup.</span></span>

7. <span data-ttu-id="79fca-500">Edite quaisquer peças necessárias adicionais em que você sabe que pode remover marcação substancial sem afetar o conteúdo, como a parte styles.</span><span class="sxs-lookup"><span data-stu-id="79fca-500">Edit any additional required parts where you know that you can remove substantial markup without affecting your content, such as the styles part.</span></span>

<span data-ttu-id="79fca-p201">Após executar as sete etapas anteriores, você provavelmente terá removido cerca de 90% a 100% da marcação que pode remover, dependendo do conteúdo. Na maioria dos casos, provavelmente esse é o máximo que você deseja cortar.</span><span class="sxs-lookup"><span data-stu-id="79fca-p201">After you've taken the preceding seven steps, you've likely cut between about 90 and 100 percent of the markup you can remove, depending on your content. In most cases, this is likely to be as far as you want to trim.</span></span>

<span data-ttu-id="79fca-503">Independentemente de você parar por aqui ou optar por se aprofundar ainda mais no conteúdo para localizar todas as linhas de marcação que pode recortar, lembre-se de que pode usar o exemplo de código referenciado anteriormente, [Word-Add-in-Get-Set-EditOpen-XML](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML), como um bloco de rascunho para testar com rapidez e facilidade a marcação editada.</span><span class="sxs-lookup"><span data-stu-id="79fca-503">Regardless of whether you leave it here or choose to delve further into your content to find every last line of markup you can cut, remember that you can use the previously-referenced code sample [Word-Add-in-Get-Set-EditOpen-XML](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML) as a scratch pad to quickly and easily test your edited markup.</span></span>


> [!TIP]
> <span data-ttu-id="79fca-p202">Se você atualizar um trecho do Office Open XML em uma solução existente durante o desenvolvimento, limpe arquivos de Internet temporários antes de executar a solução novamente para atualizar o Office Open XML usado pelo código. A marcação incluída na solução em arquivos XML é armazenada no cache no computador. Claro, você pode limpar os arquivos de Internet temporários do navegador da Web padrão. Para acessar as opções da Internet e excluir essas configurações de dentro do Visual Studio 2017, no menu **Depurar**, escolha **Opções**. Em seguida, em **Ambiente**, escolha **Navegador da Web** e **Opções do Internet Explorer**.</span><span class="sxs-lookup"><span data-stu-id="79fca-p202">If you update an Office Open XML snippet in an existing solution while developing, clear temporary Internet files before you run the solution again to update the Office Open XML used by your code. Markup that's included in your solution in XML files is cached on your computer. You can, of course, clear temporary Internet files from your default web browser. To access Internet options and delete these settings from inside Visual Studio 2017, on the  **Debug** menu, choose **Options**. Then, under  **Environment**, choose  **Web Browser** and then choose **Internet Explorer Options**.</span></span>


## <a name="creating-an-add-in-for-both-template-and-stand-alone-use"></a><span data-ttu-id="79fca-509">Criação de um suplemento para o modelo e para uso autônomo</span><span class="sxs-lookup"><span data-stu-id="79fca-509">Creating an add-in for both template and stand-alone use</span></span>


<span data-ttu-id="79fca-p203">Neste tópico, você viu vários exemplos do que pode fazer com o Office Open XML em suplementos. Vimos uma ampla variedade de exemplos de tipo de conteúdo avançado que você pode inserir em documentos usando o tipo de coerção do Office Open XML, juntamente com os métodos de JavaScript para inserir o conteúdo na seleção ou em um local específico (associado).</span><span class="sxs-lookup"><span data-stu-id="79fca-p203">In this topic, you've seen several examples of what you can do with Office Open XML in your add-ins for . We've looked at a wide range of rich content type examples that you can insert into documents by using the Office Open XML coercion type, together with the JavaScript methods for inserting that content at the selection or to a specified (bound) location.</span></span>

<span data-ttu-id="79fca-p204">Portanto, o que mais você precisa saber se estiver criando o suplemento para uso autônomo (ou seja, inserido da Loja ou em um local de servidor proprietário) e para uso em um modelo pré-criado projetado para funcionar com o suplemento? A resposta pode ser que você já sabe tudo o que precisa saber.</span><span class="sxs-lookup"><span data-stu-id="79fca-p204">So, what else do you need to know if you're creating your add-in both for stand-alone use (that is, inserted from the Store or a proprietary server location) and for use in a pre-created template that's designed to work with your add-in? The answer might be that you already know all you need.</span></span>

<span data-ttu-id="79fca-p205">A marcação de determinado tipo de conteúdo e os métodos para inseri-la são os mesmos, quer o suplemento seja projetado como autônomo, quer seja para uso com um modelo. Se você estiver usando modelos projetados para funcionar com o suplemento, verifique se o JavaScript inclui retornos de chamada que levam em conta cenários em que o conteúdo referenciado pode já existir no documento (conforme demonstrado no exemplo de associação mostrado na seção [Adicionar uma associação a um controle de conteúdo nomeado](#add-and-bind-to-a-named-content-control)).</span><span class="sxs-lookup"><span data-stu-id="79fca-p205">The markup for a given content type and methods for inserting it are the same whether your add-in is designed to stand-alone or work with a template. If you are using templates designed to work with your add-in, just be sure that your JavaScript includes callbacks that account for scenarios where referenced content might already exist in the document (such as demonstrated in the binding example shown in the section [Add and bind to a named content control](#add-and-bind-to-a-named-content-control)).</span></span>

<span data-ttu-id="79fca-p206">Ao usar modelos com o aplicativo, se o suplemento será residente no modelo no momento em que o usuário criou o documento ou se o suplemento inserirá um modelo, também convém incorporar outros elementos da API para ajudá-lo a criar uma experiência mais robusta e interativa. Por exemplo, convém incluir a identificação de dados em uma parte customXML que você pode usar para determinar o tipo de modelo para oferecer opções específicas de modelo para o usuário. Para saber mais sobre como trabalhar com XML personalizado em suplementos, confira os recursos adicionais a seguir.</span><span class="sxs-lookup"><span data-stu-id="79fca-p206">When using templates with your app, whether the add-in will be resident in the template at the time that the user created the document or the add-in will be inserting a template, you might also want to incorporate other elements of the API to help you create a more robust, interactive experience. For example, you may want to include identifying data in a customXML part that you can use to determine the template type in order to provide template-specific options to the user. To learn more about how to work with custom XML in your add-ins, see the additional resources that follow.</span></span>


## <a name="see-also"></a><span data-ttu-id="79fca-519">Confira também</span><span class="sxs-lookup"><span data-stu-id="79fca-519">See also</span></span>

- [<span data-ttu-id="79fca-520">API JavaScript para Office</span><span class="sxs-lookup"><span data-stu-id="79fca-520">JavaScript API for Office </span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- <span data-ttu-id="79fca-521">[Padrão ECMA-376: Formatos do Office Open XML](https://www.ecma-international.org/publications/standards/Ecma-376.htm) (acesse a referência de linguagem completa e a documentação relacionada do Open XML aqui)</span><span class="sxs-lookup"><span data-stu-id="79fca-521">[Standard ECMA-376: Office Open XML File Formats](https://www.ecma-international.org/publications/standards/Ecma-376.htm) (access the complete language reference and related documentation on Open XML here)</span></span>
- [<span data-ttu-id="79fca-522">Como explorar a API JavaScript para Office: associação de dados e partes XML personalizadas</span><span class="sxs-lookup"><span data-stu-id="79fca-522">Exploring the JavaScript API for Office: Data Binding and Custom XML Parts</span></span>](https://msdn.microsoft.com/magazine/dn166930.aspx)
