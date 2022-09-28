---
title: Diretrizes de ícone de estilo novo para suplementos do Office
description: Diretrizes para usar ícones de estilo Fresh em Suplementos do Office.
ms.date: 07/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: 261f684648e8fb57a3aa291b785b33e511f83865
ms.sourcegitcommit: 05be1086deb2527c6c6ff3eafcef9d7ed90922ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/28/2022
ms.locfileid: "68092942"
---
# <a name="fresh-style-icon-guidelines-for-office-add-ins"></a>Diretrizes de ícone de estilo novo para suplementos do Office

As versões do Office 2013+ (sem assinatura) do Office usam a iconografia de estilo Fresh da Microsoft. Se você preferir que seus ícones correspondam ao estilo Monoline do Microsoft 365, consulte as diretrizes de ícone de estilo [Monoline para Suplementos do Office](add-in-icons-monoline.md).

## <a name="office-fresh-visual-style"></a>Estilo visual do Office Fresh

Os ícones Fresh incluem apenas elementos comunicativos essenciais. Elementos não essenciais, como perspectiva, gradientes e uma fonte de luz, foram removidos. Os ícones simplificados suportam a análise mais rápida de comandos e controles. Siga esse estilo para se ajustar melhor aos clientes sem assinatura do Office.

## <a name="best-practices"></a>Práticas recomendadas

Siga estas diretrizes ao criar seus ícones.

|Fazer|Não fazer|
|:---|:---|
|Mantenha os visuais simples e claros, concentrando-se nos principais elementos da comunicação.| Não usar artefatos que façam com que o ícone pareça confuso.|
|Usar a linguagem de ícones do Office para representar comportamentos ou conceitos.|Não reutilize os glifos do Fabric Core para comandos de suplemento na faixa de opções do aplicativo do Office ou menus contextuais. Os ícones do Fabric Core são estilicamente diferentes e não corresponderão.|
|Reutilizar metáforas visuais comuns do Office, como o pincel para formatar ou a lupa para localizar.|Não reutilizar metáforas visuais para comandos diferentes. Usar o mesmo ícone para conceitos e comportamentos diferentes pode causar confusão. |
|Redesenhar os ícones para deixá-los pequenos ou maiores. Dedicar um tempo para redesenhar recortes, cantos e bordas arredondadas para maximizar a clareza da linha. |Não redimensionar os ícones reduzindo-os ou aumentando-os. Isso pode levar a uma baixa qualidade visual e a ações confusas. Os ícones complexos criados em um tamanho maior podem perder clareza ao ser redimensionados para ficar menores sem um redesenho. |
|Use a white fill for accessibility. Most objects in your icons will require a white background to be legible across Office UI themes and in high-contrast modes.  |Evite confiar no seu logotipo ou marca para comunicar o que um comando de suplemento faz. Nem sempre é possível reconhecer as marcas em ícones menores e quando os modificadores são aplicados. Marcas de marca geralmente entram em conflito com os estilos de ícone da faixa de opções do aplicativo do Office e podem competir pela atenção do usuário em um ambiente saturado. |
|Usar o formato PNG com uma tela de fundo transparente. |*Nenhum.*|
|Evitar usar conteúdo localizável nos ícones, como caracteres tipográficos, indicações de parágrafos e pontos de interrogação. |*Nenhum.*|

## <a name="icon-size-recommendations-and-requirements"></a>Recomendações e requisitos de tamanho de ícone

Os ícones da área de trabalho do Office são imagens bitmap. Os tamanhos diferentes serão renderizados, dependendo do modo de toque e da configuração de DPI do usuário. Inclua todos os oito tamanhos com suporte para criar a melhor experiência para todas as resoluções e contextos com suporte. A seguir estão os tamanhos com suporte: três são necessários.

- 16 px (obrigatório)
- 20 px
- 24 px
- 32 px (obrigatório)
- 40 px
- 48 px
- 64 px (recomendado, melhor para Mac)
- 80 px (obrigatório)

> [!IMPORTANT]
> Para obter uma imagem que seja o ícone representativo do suplemento, consulte Criar [listagem efetiva no AppSource](/office/dev/store/create-effective-office-store-listings#create-an-icon-for-your-add-in) e no Office para obter tamanho e outros requisitos.

Não se esqueça de redesenhar seus ícones para cada tamanho em vez de reduzi-los para que caibam.

![Ilustração da recomendação para redesenhar ícones por tamanho em vez de reduzir ícones. Por exemplo, talvez seja necessário usar menos elementos em um ícone pequeno em vez de apenas reduzir verticalmente uma imagem maior.](../images/icon-resizing.png)

## <a name="icon-anatomy-and-layout"></a>Anatomia e layout do ícone

Os ícones do Office normalmente são compostos por um elemento base com ação e modificadores conceituais sobrepostos. Os modificadores de ação representam conceitos como adicionar, abrir, novo ou fechar. Os modificadores conceituais representam status, alteração ou uma descrição do ícone.

To create commands that align with the Office UI, follow layout guidelines for the base element and modifiers. This ensures that your commands look professional and that your customers will trust your add-in. If you make exceptions to these guidelines, do so intentionally.

A imagem a seguir mostra o layout de elementos básicos e modificadores em um ícone do Office.

![Diagrama mostrando um elemento base de ícone no centro com um modificador no canto inferior direito e um modificador de ação no canto superior esquerdo.](../images/icon-layouts.png)

- Elementos básicos centrais no quadro do pixel com preenchimento todo vazio.
- Coloque modificadores de ação na parte superior esquerda.
- Coloque modificadores conceituais no canto inferior direito.
- Limite o número de elementos em seus ícones. A 32 px, limite o número de modificadores a um máximo de dois. Em 16 px, limite o número de modificadores a um.

### <a name="base-element-padding"></a>Preenchimento do elemento básico

Coloque elementos básicos com tamanhos consistentes. Se os elementos básicos não puderem ser centralizados no quadro, alinhe-os no canto superior esquerdo, deixando os pixels extras na parte inferior direita. Para obter melhores resultados, aplique as diretrizes de preenchimento listadas na tabela na seção a seguir.

### <a name="modifiers"></a>Modificadores

Todos os modificadores devem ter um recorte transparente de 1 px entre cada elemento, incluindo a tela de fundo. Os elementos não devem se sobrepor diretamente. Crie um espaço em branco entre as regras e as bordas. Os modificadores podem variar um pouco de tamanho, mas use essas dimensões como ponto de partida.

|Tamanho do ícone|Preenchimento em torno do elemento básico|Tamanho do modificador|
|:---|:---|:---|
|16 px|0|9 px|
|20 px|1px|10 px|
|24 px|1px|12 px|
|32 px|2px|14 px|
|40 px|2px|20 px|
|48 px|3px|22 px|
|64 px|5px|29 px|
|80 px|5px|38 px|

## <a name="icon-colors"></a>Cores do ícone

> [!NOTE]
> Estas diretrizes de cor são destinadas a ícones da faixa de opções usados em [Comandos do suplemento](add-in-commands.md). Esses ícones não são renderizados com a interface do usuário fluente e a paleta de cores é diferente da paleta descrita no [Microsoft UI Fabric | Cores | Compartilhado, compartilhado](https://fluentfabric.azurewebsites.net/#/color/shared).

Os ícones do Office têm uma paleta de cores limitada. Use as cores listadas na tabela a seguir para garantir uma integração perfeita com a interface de usuário do Office. Aplique as diretrizes a seguir ao uso de cores.

- Use color to communicate meaning rather than for embellishment. It should highlight or emphasize an action, status, or an element that explicitly differentiates the mark.
- Se possível, use somente uma cor além do cinza. Limite as cores adicionais a no máximo duas.
- As cores devem ter uma aparência consistente em todos os tamanhos de ícone. Os ícones do Office têm paletas de cores um pouco diferentes para tamanhos de ícones diferentes. 16 px e ícones menores são um pouco mais escuros e mais vibrantes do que 32 px e ícones maiores. Sem esses ajustes sutis, as cores parecem variar entre os tamanhos.

|Nome da cor|RGB|Hex|Cor|Categoria|
|:---|:---|:---|:---|:---|
|Texto Cinza (80)|80, 80, 80|#505050| ![Cor cinza 80 para texto.](../images/color-text-gray-80.png) |Texto|
|Texto Cinza (95)|95, 95, 95|#5F5F5F| ![Cor cinza 95 para texto.](../images/color-text-gray-95.png) |Texto|
|Texto Cinza (105)|105, 105, 105|#696969| ![Cor cinza 105 para texto.](../images/color-text-gray-105.png) |Texto|
|Cinza Escuro 32|128, 128, 128|#808080| ![Cor cinza escuro para 32 px e maior.](../images/color-dark-gray-32.png) |32 px e superior|
|Cinza Médio 32|158, 158, 158|#9E9E9E| ![Cor cinza média para 32 px e maior.](../images/color-medium-gray-32.png) |32 px e superior|
|Cinza Claro TODO|179, 179, 179|#B3B3B3| ![Cor cinza claro para todos os tamanhos de imagem.](../images/color-light-gray-all.png) |Todos os tamanhos|
|Cinza Escuro 16|114, 114, 114|#727272| ![Cor cinza escuro para 16 px e menor.](../images/color-dark-gray-16.png) |16 px e inferiores|
|Cinza Médio 16|144, 144, 144|#909090| ![Cor cinza média para 16 px e menor.](../images/color-medium-gray-16.png) |16 e abaixo|
|Azul 32|77, 130, 184|#4d82B8| ![Cor azul para 32 px e maior.](../images/color-blue-32.png) |32 px e superior|
|Azul 16|74, 125, 177|#4A7DB1| ![Cor azul para 16 px e menor.](../images/color-blue-16.png) |16 px e inferiores|
|Amarelo TODO|234, 194, 130|#EAC282| ![Cor amarela para todos os tamanhos de imagem.](../images/color-yellow-all.png) |Todos os tamanhos|
|Laranja 32|231, 142, 70|#E78E46| ![Cor laranja para 32 px e maior.](../images/color-orange-32.png) |32 px e superior|
|Laranja 16|227, 142, 70|#E3751C| ![Cor laranja para 16 px e menor.](../images/color-orange-16.png) |16 px e inferiores|
|Rosa TODO|230, 132, 151|#E68497| ![Cor rosa para todos os tamanhos de imagem.](../images/color-pink-all.png) |Todos os tamanhos|
|Verde 32|118, 167, 151|#76A797| ![Cor verde para 32 px e maior.](../images/color-green-32.png) |32 px e superior|
|Verde 16|104, 164, 144|#68A490| ![Cor verde para 16 px e menor.](../images/color-green-16.png) |16 px e inferiores|
|Vermelho 32|216, 99, 68|#D86344| ![Cor vermelha para 32 px e maior.](../images/color-red-32.png) |32 px e superior|
|Vermelho 16|214, 85, 50|#D65532| ![Cor vermelha para 16 px e menor.](../images/color-red-16.png) |16 px e inferiores|
|Roxo 32|152, 104, 185|#9868B9| ![Cor roxa para 32 px e maior.](../images/color-purple-32.png) |32 px e superior|
|Roxo 16|137, 89, 171|#8959AB| ![Cor roxa para 16 px e menor.](../images/color-purple-16.png) |16 px e inferiores|

## <a name="icons-in-high-contrast-modes"></a>Ícones em modos de alto contraste

Os ícones do Office foram projetados para renderizar bem em modos de alto contraste. Elementos de primeiro plano são bem diferenciados dos planos de fundo para maximizar a legibilidade e habilitar a recoloração. Nos modos de alto contraste, o Office recolore qualquer pixel do seu ícone com um valor vermelho, verde ou azul menor que 190 para totalmente preto. Todos os outros pixels ficam na cor branca. Em outras palavras, cada canal RGB é avaliado onde, os valores de 0 a 189 ficam pretos e os valores de 190 a 255 ficam brancos. Outros temas de alto contraste fazem a recoloração usando o mesmo limite de valor de 190, mas com regras diferentes. Por exemplo, o tema de branco de alto contraste recolore todos pixels maiores que 190 para opaco, mas todos os outros pixels para transparente. Aplique as diretrizes a seguir para maximizar a legibilidade em configurações de alto contraste.

- Vise diferenciar elementos de primeiro plano e de plano de fundo ao longo do limite de valor de 190.
- Siga os estilos visuais dos ícones do Office.
- Use cores da nossa paleta de ícones.
- Evite o uso de gradientes.
- Evite blocos grandes de cores com valores similares.

## <a name="see-also"></a>Confira também

- [Elemento de manifesto do ícone](/javascript/api/manifest/icon)
- [Elemento de manifesto IconUrl](/javascript/api/manifest/iconurl)
- [Elemento de manifesto HighResolutionIconUrl](/javascript/api/manifest/highresolutioniconurl)
- [Criar um ícone para o seu suplemento](/office/dev/store/create-effective-office-store-listings#create-an-icon-for-your-add-in)
