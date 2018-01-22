# <a name="design-icons-for-add-in-commands"></a>Desenvolva ícones para comandos de suplemento

Os [Comandos de suplementos](add-in-commands.md) adicionam botões, texto e ícones à interface do usuário do Office. Os botões de comando de suplemento devem fornecer ícones significativos e rótulos que identifiquem claramente a ação que o usuário está realizando ao usar um comando. Este artigo fornece diretrizes de estilo e produção que ajudam você a desenvolver ícones que se integrem perfeitamente ao Office. 

## <a name="office-icon-design-principles"></a>Princípios de design de ícones do Office

A versão Office 2013 de clientes de área de trabalho do Office conta com uma iconografia atualizada. A mudança estilística de substituição é a redução. Os novos ícones contêm apenas elementos de comunicação essenciais. Elementos não essenciais, como perspectiva, gradientes e uma fonte de luz, foram removidos. Os ícones simplificados suportam a análise mais rápida de comandos e controles. Siga esse estilo para ter uma melhor integração com o Office.

Os ícones do Office são baseados nas seguintes princípios de design: 

- Interpretação moderna do conjunto de ícones do Office 
- Novo, porém reconhecível  
- Simples, claro e direto 

A imagem a seguir mostra ícones que aplicam os princípios modernos de design.

![Imagem mostrando ícones antigos do Office e a interpretação moderna e atualizada dos ícones](../images/icons_image.PNG)

## <a name="icon-guidelines"></a>Diretrizes de ícones
Siga estas diretrizes ao criar seus ícones: 

- Mantenha uma grade de 1 px e use uma ferramenta de edição bitmap para obter melhores resultados.  
- Redesenhe, não redimensione. À medida que você redimensiona seus ícones para tamanhos maiores ou menores, reserve um tempo para redesenhar os recortes, os cantos e as bordas arredondadas para maximizar a definição da linha. 
- Remova artefatos que façam com que o ícone pareça confuso.
- Não reutilize ícones do Office UI Fabric na faixa de opções do Office ou no menu contextual. Os ícones do Fabric são estilisticamente diferentes e não serão compatíveis. 
- Evite confiar no seu logotipo ou marca para comunicar o que um comando de suplemento faz. Nem sempre é possível reconhecer as marcas em ícones menores e quando os modificadores são aplicados. As marcas geralmente entram em conflito com estilos de ícone da faixa de opções e podem competir pela atenção do usuário em um ambiente saturado.
- Use um preenchimento branco para acessibilidade. A maioria dos objetos em seus ícones exigirá um fundo branco para ser legível nos temas da interface do usuário do Office e nos modos de alto contraste.  
- Use o formato PNG com uma tela de fundo transparente. 
- Evite usar conteúdo localizável em seus ícones, como caracteres tipográficos, indicações de parágrafos e pontos de interrogação. 
- Não reutilize metáforas visuais para comandos diferentes. Usar o mesmo ícone para ações diferentes pode causar confusão. 
- Faça com que os rótulos dos seus botões sejam claros e concisos. Use uma combinação de informações visuais e textuais para transmitir o significado. 


## <a name="icon-size-recommendations-and-requirements"></a>Recomendações e requisitos de tamanho de ícone

Os ícones de área de trabalho do Office 2016 são imagens bitmap. Tamanhos diferentes serão renderizados, dependendo do modo de toque e da configuração de DPI do usuário. Inclua todos os oito tamanhos com suporte para criar a melhor experiência para todas as resoluções e contextos com suporte. Estes são os tamanhos compatíveis (três são obrigatórios):

- 16 px (obrigatório)
- 20 px
- 24 px
- 32 px (obrigatório)
- 40 px
- 48 px
- 64 px (recomendado, melhor para Mac)
- 80 px (obrigatório)  

Não se esqueça de redesenhar seus ícones para cada tamanho em vez de reduzi-los para que caibam.

![Ilustração que mostra a recomendação de redimensionar os ícones em vez de reduzi-los](../images/icon_resizing.png)

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

>**Note:** At DPI settings of 150% or greater, the icon does not get swapped out for a larger size when Touch mode is engaged. At DPI settings greater than 250%, Touch mode is turned off by default.

The following table lists the locations for certain icon sizes.

|Location|100% DPI|200% DPI|250% DPI|
|:-------|:-------|:-------|:-------|
|Small ribbon button|16px|32px|40px|
|Contextual menu|16px|32px|40px|
|Quick access toolbar (QAT)|16px|32px|40px|
|Large ribbon icon|32px|64px|80px|

-->

## <a name="icon-anatomy-and-layout"></a>Anatomia e layout do ícone

Os ícones do Office são geralmente compostos por um elemento básico com modificadores conceituais e de ação sobrepostos. Os modificadores de ação representam conceitos como adicionar, abrir, novo ou fechar. Os modificadores conceituais representam status, alteração ou uma descrição do ícone. 

Para criar comandos que se alinhem à interface do usuário do Office, siga as diretrizes de layout para o elemento básico e os modificadores. Isso garante que seus comandos tenham uma aparência profissional e seus clientes confiem no seu suplemento. Se você fizer exceções a essas diretrizes, faça intencionalmente.

A imagem a seguir mostra o layout de elementos básicos e modificadores em um ícone do Office.

![Imagem mostrando o elemento básico de um ícone no centro com um modificador no canto inferior direito e um modificador de ação no canto superior esquerdo](../images/icon_layout.PNG)

- Elementos básicos centrais no quadro do pixel com preenchimento todo vazio.
- Coloque modificadores de ação na parte superior esquerda. 
- Coloque modificadores conceituais no canto inferior direito.
- Limite o número de elementos em seus ícones. Com 32 px, limite o número de modificadores a no máximo dois. Com 16 px, limite o número de modificadores a um.

Coloque elementos básicos com tamanhos consistentes. Se os elementos básicos não puderem ser centralizados no quadro, alinhe-os no canto superior esquerdo, deixando os pixels extras na parte inferior direita. Para melhores resultados, aplique as diretrizes de preenchimento listadas na tabela a seguir.

|**Tamanho do ícone**|**Preenchimento em torno do elemento básico**|
|:---|:---|
|16px|0|
|20px|1px|
|24px|1px|
|32px|2px|
|40px|2px|
|48px|3px|
|64px|5px|
|80px|5px|

Todos os modificadores devem ter um recorte transparente de 1 px entre cada elemento, incluindo a tela de fundo. Os elementos não devem se sobrepor diretamente. Crie um espaço em branco entre as regras e as bordas. Os modificadores podem variar um pouco de tamanho, mas use essas dimensões como ponto de partida.

|**Tamanho do ícone**|**Tamanho do modificador**|
|:---|:---|
|16px|9px|
|20px|10px|
|24px|12px|
|32px|14px|
|40px|20px|
|48px|22px|
|64px|29px|
|80px|38px|

## <a name="icon-colors"></a>Cores do ícone

Os ícones do Office têm uma paleta de cores limitada. Use as cores listadas na tabela a seguir para garantir uma integração perfeita com a interface de usuário do Office. Aplique as seguintes diretrizes para o uso de cor: 

- Use cor para transmitir significado, não como enfeite. Ela deve destacar ou enfatizar uma ação, status ou um elemento que diferencie explicitamente a marca.  
- Se possível, use somente uma cor além do cinza. Limite as cores adicionais a no máximo duas.
- As cores devem ter uma aparência consistente em todos os tamanhos de ícone. Os ícones do Office têm paletas de cores um pouco diferentes para tamanhos de ícones diferentes. Ícones com 16 px e menores são um pouco mais escuros e mais vibrantes do que os ícones de 32 px e maiores. Sem esses ajustes sutis, as cores parecem variar entre os tamanhos.   

|**Nome da cor**|**RGB**|**Hex**|**Cor**|**Categoria**|
|:---|:---|:---|:---|:---|
|Texto Cinza (80)|80, 80, 80|#505050|![Imagem colorida texto cinza 80](../images/textGray_80.gif)|Texto|
|Texto Cinza (95)|95, 95, 95|#5F5F5F|![Imagem colorida texto cinza 95](../images/textGray_95.gif)|Texto|
|Texto Cinza (105)|105, 105, 105|#696969|![Imagem colorida texto cinza 105](../images/textGray_105.gif)|Texto|
|Cinza Escuro 32|128, 128, 128|#808080|![Imagem colorida cinza escuro 32](../images/darkGray_32.gif)|32 e acima|
|Cinza Médio 32|158, 158, 158|#9E9E9E|![Imagem colorida cinza médio 32](../images/mediumGray_32.gif)|32 e acima|
|Cinza Claro TODO|179, 179, 179|#B3B3B3|![Imagem colorida cinza claro todo](../images/lightGray_all.gif)|Todos os tamanhos|
|Cinza Escuro 16|114, 114, 114|#727272|![Imagem colorida cinza escuro 16](../images/darkGray_16.gif)|16 e abaixo|
|Cinza Médio 16|144, 144, 144|#909090|![Imagem colorida cinza médio 16](../images/mediumGray_16.gif)|16 e abaixo|
|Azul 32|77, 130, 184|#4d82B8|![Imagem colorida azul 32](../images/blue_32.gif)|32 e acima|
|Azul 16|74, 125, 177|#4A7DB1|![Imagem colorida azul 16](../images/blue_16.gif)|16 e abaixo|
|Amarelo TODO|234, 194, 130|#EAC282|![Imagem colorida amarelo todo](../images/yellow_all.gif)|Todos os tamanhos|
|Laranja 32|231, 142, 70|#E78E46|![Imagem colorida laranja 32](../images/orange_32.gif)|32 e acima|
|Laranja 16|227, 142, 70|#E3751C|![Imagem colorida laranja 16](../images/orange_16.gif)|16 e abaixo|
|Rosa TODO|230, 132, 151|#E68497|![Imagem colorida rosa todo](../images/pink_all.gif)|Todos os tamanhos|
|Verde 32|118, 167, 151|#76A797|![Imagem colorida verde 32](../images/green_32.gif)|32 e acima|
|Verde 16|104, 164, 144|#68A490|![Imagem colorida verde 16](../images/green_16.gif)|16 e abaixo|
|Vermelho 32|216, 99, 68|#D86344|![Imagem colorida vermelho 32](../images/red_32.gif)|32 e acima|
|Vermelho 16|214, 85, 50|#D65532|![Imagem colorida vermelho 16](../images/red_16.gif)|16 e abaixo|
|Roxo 32|152, 104, 185|#9868B9|![Imagem colorida roxo 32](../images/purple_32.gif)|32 e acima|
|Roxo 16|137, 89, 171|#8959AB|![Imagem colorida roxo 16](../images/purple_16.gif)|16 e abaixo|

## <a name="icons-in-high-contrast-modes"></a>Ícones em modos de alto contraste

Os ícones do Office foram projetados para renderizar bem em modos de alto contraste. Elementos de primeiro plano são bem diferenciados dos planos de fundo para maximizar a legibilidade e habilitar a recoloração. Nos modos de alto contraste, o Office recolore qualquer pixel do seu ícone com um valor vermelho, verde ou azul menor que 190 para totalmente preto. Todos os outros pixels ficam na cor branca. Em outras palavras, cada canal RGB é avaliado onde, os valores de 0 a 189 ficam pretos e os valores de 190 a 255 ficam brancos. Outros temas de alto contraste fazem a recoloração usando o mesmo limite de valor de 190, mas com regras diferentes. Por exemplo, o tema de branco de alto contraste recolore todos pixels maiores que 190 para opaco, mas todos os outros pixels para transparente. Aplique as seguintes diretrizes para maximizar a legibilidade em configurações de alto contraste:

- Vise diferenciar elementos de primeiro plano e de plano de fundo ao longo do limite de valor de 190.
- Siga os estilos visuais dos ícones do Office.
- Use cores da nossa paleta de ícones.
- Evite o uso de gradientes.
- Evite blocos grandes de cores com valores similares.

## <a name="additional-resources"></a>Recursos adicionais

- [Práticas recomendadas de desenvolvimento de suplementos](../overview/add-in-development-best-practices.md)
- [Comandos de suplemento para Excel, Word e PowerPoint](../design/add-in-commands.md)
