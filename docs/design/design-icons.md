---
title: Desenvolver ?cones para comandos de suplemento
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: a3dc7837bdc95df9576a5fc4a6c1840e64afacb6
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="design-icons-for-add-in-commands"></a>Desenvolver ?cones para comandos de suplemento

Os [Comandos de suplementos](add-in-commands.md) adicionam bot?es, texto e ?cones ? interface do usu?rio do Office. Os bot?es de comando de suplemento devem fornecer ?cones significativos e r?tulos que identifiquem claramente a a??o que o usu?rio est? realizando ao usar um comando. Este artigo fornece diretrizes de estilo e produ??o que ajudam voc? a desenvolver ?cones que se integrem perfeitamente ao Office. 

## <a name="office-icon-design-principles"></a>Princ?pios de design de ?cones do Office

A vers?o Office 2013 de clientes de ?rea de trabalho do Office conta com uma iconografia atualizada. A mudan?a estil?stica de substitui??o ? a redu??o. Os novos ?cones cont?m apenas elementos de comunica??o essenciais. Elementos n?o essenciais, como perspectiva, gradientes e uma fonte de luz, foram removidos. Os ?cones simplificados suportam a an?lise mais r?pida de comandos e controles. Siga esse estilo para ter uma melhor integra??o com o Office.

Os ?cones do Office s?o baseados nas seguintes princ?pios de design: 

- Interpreta??o moderna do conjunto de ?cones do Office 
- Novo, por?m reconhec?vel  
- Simples, claro e direto 

A imagem a seguir mostra ?cones que aplicam os princ?pios modernos de design.

![Imagem mostrando ?cones antigos do Office e a interpreta??o moderna e atualizada dos ?cones](../images/icons-images.png)

## <a name="icon-guidelines"></a>Diretrizes de ?cones
Siga estas diretrizes ao criar seus ?cones: 

- Mantenha uma grade de 1 px e use uma ferramenta de edi??o bitmap para obter melhores resultados.  
- Redesenhe, n?o redimensione. ? medida que voc? redimensiona seus ?cones para tamanhos maiores ou menores, reserve um tempo para redesenhar os recortes, os cantos e as bordas arredondadas para maximizar a defini??o da linha. 
- Remova artefatos que fa?am com que o ?cone pare?a confuso.
- N?o reutilize ?cones do Office UI Fabric na faixa de op??es do Office ou no menu contextual. Os ?cones do Fabric s?o estilisticamente diferentes e n?o ser?o compat?veis. 
- Evite confiar no seu logotipo ou marca para comunicar o que um comando de suplemento faz. Nem sempre ? poss?vel reconhecer as marcas em ?cones menores e quando os modificadores s?o aplicados. As marcas geralmente entram em conflito com estilos de ?cone da faixa de op??es e podem competir pela aten??o do usu?rio em um ambiente saturado.
- Use um preenchimento branco para acessibilidade. A maioria dos objetos em seus ?cones exigir? um fundo branco para ser leg?vel nos temas da interface do usu?rio do Office e nos modos de alto contraste.  
- Use o formato PNG com uma tela de fundo transparente. 
- Evite usar conte?do localiz?vel em seus ?cones, como caracteres tipogr?ficos, indica??es de par?grafos e pontos de interroga??o. 
- N?o reutilize met?foras visuais para comandos diferentes. Usar o mesmo ?cone para a??es diferentes pode causar confus?o. 
- Fa?a com que os r?tulos dos seus bot?es sejam claros e concisos. Use uma combina??o de informa??es visuais e textuais para transmitir o significado. 


## <a name="icon-size-recommendations-and-requirements"></a>Recomenda??es e requisitos de tamanho de ?cone

Os ?cones de ?rea de trabalho do Office 2016 s?o imagens bitmap. Tamanhos diferentes ser?o renderizados, dependendo do modo de toque e da configura??o de DPI do usu?rio. Inclua todos os oito tamanhos com suporte para criar a melhor experi?ncia para todas as resolu??es e contextos com suporte. Estes s?o os tamanhos compat?veis (tr?s s?o obrigat?rios):

- 16 px (obrigat?rio)
- 20 px
- 24 px
- 32 px (obrigat?rio)
- 40 px
- 48 px
- 64 px (recomendado, melhor para Mac)
- 80 px (obrigat?rio)  

N?o se esque?a de redesenhar seus ?cones para cada tamanho em vez de reduzi-los para que caibam.

![Ilustra??o que mostra a recomenda??o de redimensionar os ?cones em vez de reduzi-los](../images/icon-resizing.png)

<!--
The following table shows the icon sizes that render for different modes at different DPI settings.

|DPI |**Small**||**Medium**||**Large**||**Extra large**|
|:---|:---|:---|:---|:---|:---|:---|:---|
|    |**Mouse**|**Touch**|**Mouse**|**Touch**|**Mouse**|**Touch**|-|
|100%|16px|20px|24px||32px|40px|48px|
|125%|20px|24px|||40px|48px|60px|
|150%|24px|24px|36px||48px|48px|72px|
|200%|32px|40px|48px||64px|80px|96px|
|250%|40px||||80px||120px|
|300%|48px||||96px||144px

> [!NOTE]
> At DPI settings of 150% or greater, the icon does not get swapped out for a larger size when Touch mode is engaged. At DPI settings greater than 250%, Touch mode is turned off by default.

The following table lists the locations for certain icon sizes.

|Location|100% DPI|200% DPI|250% DPI|
|:-------|:-------|:-------|:-------|
|Small ribbon button|16px|32px|40px|
|Contextual menu|16px|32px|40px|
|Quick access toolbar (QAT)|16px|32px|40px|
|Large ribbon icon|32px|64px|80px|

-->

## <a name="icon-anatomy-and-layout"></a>Anatomia e layout do ?cone

Os ?cones do Office s?o geralmente compostos por um elemento b?sico com modificadores conceituais e de a??o sobrepostos.?Os modificadores de a??o representam conceitos como adicionar, abrir, novo ou fechar. Os modificadores conceituais representam status, altera??o ou uma descri??o do ?cone. 

Para criar comandos que se alinhem ? interface do usu?rio do Office, siga as diretrizes de layout para o elemento b?sico e os modificadores. Isso garante que seus comandos tenham uma apar?ncia profissional e seus clientes confiem no seu suplemento. Se voc? fizer exce??es a essas diretrizes, fa?a intencionalmente.

A imagem a seguir mostra o layout de elementos b?sicos e modificadores em um ?cone do Office.

![Imagem mostrando o elemento b?sico de um ?cone no centro com um modificador no canto inferior direito e um modificador de a??o no canto superior esquerdo](../images/icon-layouts.png)

- Elementos b?sicos centrais no quadro do pixel com preenchimento todo vazio.
- Coloque modificadores de a??o na parte superior esquerda. 
- Coloque modificadores conceituais no canto inferior direito.
- Limite o n?mero de elementos em seus ?cones. Com 32 px, limite o n?mero de modificadores a no m?ximo dois. Com 16 px, limite o n?mero de modificadores a um.

Coloque elementos b?sicos com tamanhos consistentes. Se os elementos b?sicos n?o puderem ser centralizados no quadro, alinhe-os no canto superior esquerdo, deixando os pixels extras na parte inferior direita. Para melhores resultados, aplique as diretrizes de preenchimento listadas na tabela a seguir.

|**Tamanho do ?cone**|**Preenchimento em torno do elemento b?sico**|
|:---|:---|
|16 px|0|
|20 px|1 px|
|24 px|1 px|
|32 px|2 px|
|40 px|2 px|
|48 px|3 px|
|64 px|5 px|
|80 px|5 px|

Todos os modificadores devem ter um recorte transparente de 1 px entre cada elemento, incluindo a tela de fundo. Os elementos n?o devem se sobrepor diretamente. Crie um espa?o em branco entre as regras e as bordas. Os modificadores podem variar um pouco de tamanho, mas use essas dimens?es como ponto de partida.

|**Tamanho do ?cone**|**Tamanho do modificador**|
|:---|:---|
|16 px|9 px|
|20 px|10 px|
|24 px|12 px|
|32 px|14 px|
|40 px|20 px|
|48 px|22 px|
|64 px|29 px|
|80 px|38 px|

## <a name="icon-colors"></a>Cores do ?cone

Os ?cones do Office t?m uma paleta de cores limitada. Use as cores listadas na tabela a seguir para garantir uma integra??o perfeita com a interface de usu?rio do Office. Aplique as seguintes diretrizes para o uso de cor: 

- Use cor para transmitir significado, n?o como enfeite. Ela deve destacar ou enfatizar uma a??o, status ou um elemento que diferencie explicitamente a marca.  
- Se poss?vel, use somente uma cor al?m do cinza. Limite as cores adicionais a no m?ximo duas.
- As cores devem ter uma apar?ncia consistente em todos os tamanhos de ?cone. Os ?cones do Office t?m paletas de cores um pouco diferentes para tamanhos de ?cones diferentes. ?cones com 16 px e menores s?o um pouco mais escuros e mais vibrantes do que os ?cones de 32 px e maiores. Sem esses ajustes sutis, as cores parecem variar entre os tamanhos.   

|**Nome da cor**|**RGB**|**Hexa**|**Cor**|**Categoria**|
|:---|:---|:---|:---|:---|
|Texto Cinza (80)|80, 80, 80|#505050| ![Imagem colorida texto cinza 80](../images/color-text-gray-80.png) |Texto|
|Texto Cinza (95)|95, 95, 95|#5F5F5F| ![Imagem colorida texto cinza 95](../images/color-text-gray-95.png) |Texto|
|Texto Cinza (105)|105, 105, 105|#696969| ![Imagem colorida texto cinza 105](../images/color-text-gray-105.png) |Texto|
|Cinza Escuro 32|128, 128, 128|#808080| ![Imagem colorida cinza escuro 32](../images/color-dark-gray-32.png) |32 e acima|
|Cinza M?dio 32|158, 158, 158|#9E9E9E| ![Imagem colorida cinza m?dio 32](../images/color-medium-gray-32.png) |32 e acima|
|Cinza Claro TODO|179, 179, 179|#B3B3B3| ![Imagem colorida cinza claro todo](../images/color-light-gray-all.png) |Todos os tamanhos|
|Cinza Escuro 16|114, 114, 114|#727272| ![Imagem colorida cinza escuro 16](../images/color-dark-gray-16.png) |16 e abaixo|
|Cinza M?dio 16|144, 144, 144|#909090| ![Imagem colorida cinza m?dio 16](../images/color-medium-gray-16.png) |16 e abaixo|
|Azul 32|77, 130, 184|#4d82B8| ![Imagem colorida azul 32](../images/color-blue-32.png) |32 e acima|
|Azul 16|74, 125, 177|#4A7DB1| ![Imagem colorida azul 16](../images/color-blue-16.png) |16 e abaixo|
|Amarelo TODO|234, 194, 130|#EAC282| ![Imagem colorida amarelo todo](../images/color-yellow-all.png) |Todos os tamanhos|
|Laranja 32|231, 142, 70|#E78E46| ![Imagem colorida laranja 32](../images/color-orange-32.png) |32 e acima|
|Laranja 16|227, 142, 70|#E3751C| ![Imagem colorida laranja 16](../images/color-orange-16.png) |16 e abaixo|
|Rosa TODO|230, 132, 151|#E68497| ![Imagem colorida rosa todo](../images/color-pink-all.png) |Todos os tamanhos|
|Verde 32|118, 167, 151|#76A797| ![Imagem colorida verde 32](../images/color-green-32.png) |32 e acima|
|Verde 16|104, 164, 144|#68A490| ![Imagem colorida verde 16](../images/color-green-16.png) |16 e abaixo|
|Vermelho 32|216, 99, 68|#D86344| ![Imagem colorida vermelho 32](../images/color-red-32.png) |32 e acima|
|Vermelho 16|214, 85, 50|#D65532| ![Imagem colorida vermelho 16](../images/color-red-16.png) |16 e abaixo|
|Roxo 32|152, 104, 185|#9868B9| ![Imagem colorida roxo 32](../images/color-purple-32.png) |32 e acima|
|Roxo 16|137, 89, 171|#8959AB| ![Imagem colorida roxo 16](../images/color-purple-16.png) |16 e abaixo|


## <a name="icons-in-high-contrast-modes"></a>?cones em modos de alto contraste

Os ?cones do Office foram projetados para renderizar bem em modos de alto contraste. Elementos de primeiro plano s?o bem diferenciados dos planos de fundo para maximizar a legibilidade e habilitar a recolora??o. Nos modos de alto contraste, o Office recolore qualquer pixel do seu ?cone com um valor vermelho, verde ou azul menor que 190 para totalmente preto. Todos os outros pixels ficam na cor branca. Em outras palavras, cada canal RGB ? avaliado onde, os valores de 0 a 189 ficam pretos e os valores de 190 a 255 ficam brancos. Outros temas de alto contraste fazem a recolora??o usando o mesmo limite de valor de 190, mas com regras diferentes. Por exemplo, o tema de branco de alto contraste recolore todos pixels maiores que 190 para opaco, mas todos os outros pixels para transparente. Aplique as seguintes diretrizes para maximizar a legibilidade em configura??es de alto contraste:

- Vise diferenciar elementos de primeiro plano e de plano de fundo ao longo do limite de valor de 190.
- Siga os estilos visuais dos ?cones do Office.
- Use cores da nossa paleta de ?cones.
- Evite o uso de gradientes.
- Evite blocos grandes de cores com valores similares.

## <a name="see-also"></a>Veja tamb?m

- [Pr?ticas recomendadas de desenvolvimento de suplementos](../concepts/add-in-development-best-practices.md)
- [Comandos de suplemento para Excel, Word e PowerPoint](../design/add-in-commands.md)
