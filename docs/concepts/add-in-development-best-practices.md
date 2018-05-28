---
title: Pr?ticas recomendadas para o desenvolvimento de suplementos do Office
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: e089b5ccbe9e8aa06a1622dba354b81bce1ddd4a
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="best-practices-for-developing-office-add-ins"></a>Pr?ticas recomendadas para o desenvolvimento de suplementos do Office

Os suplementos eficazes oferecem uma funcionalidade exclusiva e fascinante que estende os aplicativos do Office de uma maneira visualmente atraente. Para criar um excelente suplemento, ofere?a uma primeira experi?ncia envolvente para seus usu?rios, desenvolva uma experi?ncia de interface de usu?rio de alto n?vel e otimize o desempenho do seu suplemento. Aplique as pr?ticas recomendadas descritas neste artigo para criar suplementos que ajudem os usu?rios a concluir suas tarefas de forma r?pida e eficiente.

> [!NOTE]
> Caso pretenda [publicar](../publish/publish.md) o suplemento na experi?ncia do Office depois de cri?-lo, verifique se voc? est? em conformidade com as [Pol?ticas de valida??o do AppSource](https://docs.microsoft.com/en-us/office/dev/store/validation-policies). Por exemplo, para passar na valida??o, seu suplemento deve funcionar em todas as plataformas com suporte aos m?todos que voc? definir (para mais informa??es, confira a [se??o 4.12](https://docs.microsoft.com/en-us/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) e a [P?gina de hospedagem e disponibilidade de suplementos do Office](../overview/office-add-in-availability.md)). 

## <a name="provide-clear-value"></a>Fornecer um valor claro

- Crie suplementos que ajudem os usu?rios a concluir tarefas de forma r?pida e eficiente. Concentre-se nos cen?rios que fazem sentido para aplicativos do Office. Por exemplo:
 - Torne as principais tarefas de cria??o mais r?pidas e f?ceis, com menos interrup??es.
 - Habilite novos cen?rios no Office.
 - Incorpore servi?os complementares nos hosts do Office.
 - Melhore a experi?ncia do Office para aumentar a produtividade.
- Certifique-se de que o valor do seu suplemento seja claro para os usu?rios desde o princ?pio, [criando uma experi?ncia envolvente na primeira execu??o](#create-an-engaging-first-run-experience).
- Crie uma [listagem eficaz do AppSource](https://docs.microsoft.com/en-us/office/dev/store/create-effective-office-store-listings). Deixe claro quais s?o os benef?cios do seu suplemento no t?tulo e na descri??o. N?o dependa da sua marca para dizer o que seu suplemento faz.


## <a name="create-an-engaging-first-run-experience"></a>Criar uma experi?ncia envolvente na primeira execu??o

- Envolva os novos usu?rios com uma primeira experi?ncia altamente ?til e intuitiva. Observe que, mesmo depois de baixar o suplemento da loja, os usu?rios ainda est?o decidindo se v?o utiliz?-lo.

- Deixe claro quais s?o as etapas que usu?rio ter? que seguir para se envolver com seu suplemento. Use v?deos, diagramas, pain?is de pagina??o ou outros recursos para atrair usu?rios.

- Reforce a proposta de valor do seu suplemento no in?cio, em vez de apenas pedir que seus usu?rios entrem.

- Forne?a uma interface do usu?rio informativa e torne sua interface do usu?rio pessoal.

   ![Uma captura de tela que mostra um painel de tarefas de suplemento com etapas de introdu??o ao lado de um suplemento sem etapas de introdu??o](../images/contoso-part-catalog-do-dont.png)

- Se seu suplemento de conte?do estiver vinculado a dados no documento do usu?rio, inclua exemplos de dados ou um modelo para mostrar aos usu?rios o formato de dados a ser usado.

   ![Uma captura de tela que mostra um suplemento de conte?do com dados ao lado de um suplemento de conte?do sem dados](../images/add-in-title.png)

- Ofere?a [avalia??es gratuitas](https://docs.microsoft.com/en-us/office/dev/store/decide-on-a-pricing-model#office-store-pricing-options). Caso o suplemento exija uma assinatura, disponibilize algumas funcionalidades sem a necessidade da assinatura.

- Simplifique o processo de inscri??o. Preencha automaticamente as informa??es (email, nome de exibi??o) e ignore as verifica??es de email.

- Evite os pop-ups. Se voc? tiver de us?-los, oriente o usu?rio sobre como habilitar o seu pop-up.

- Use a [autentica??o de logon ?nico (SSO)](https://docs.microsoft.com/en-us/outlook/add-ins/authenticate-a-user-with-an-identity-token).

Para modelos que ilustram padr?es que podem ser aplicados enquanto voc? desenvolve sua experi?ncia na primeira execu??o, consulte [padr?es de design UX para Suplementos do Office](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code).

## <a name="use-add-in-commands"></a>Usar comandos de suplemento

- Fornece ao suplemento pontos de entrada relevantes da interface do usu?rio usando os comandos do suplemento. Confira mais detalhes, inclusive as pr?ticas recomendadas de design, nos [comandos de suplemento](../design/add-in-commands.md).

## <a name="apply-ux-design-principles"></a>Aplicar os princ?pios de design de UX

- Assegure-se de que a apar?ncia e a funcionalidade de seus suplementos complementam a experi?ncia do Office. Use o [Office UI Fabric](https://dev.office.com/fabric).

- Favore?a o conte?do atrav?s do Chrome. Evite elementos de interface do usu?rio sup?rfluos que n?o agregam valor ? experi?ncia do usu?rio.

- Mantenha os usu?rios no controle. Verifique se os usu?rios compreenderam as decis?es importantes e podem reverter facilmente as a??es realizadas pelo suplemento.

- Use uma identidade visual para inspirar confian?a e orientar os usu?rios. N?o use o recurso de identidade visual para sobrecarregar ou enviar an?ncios aos usu?rios.

- Evite a necessidade de rolagem. Otimize para a resolu??o 1366 x 768.

- N?o inclua imagens n?o licenciadas.

- Use uma [linguagem clara e simples](../design/add-in-design-guidelines.md#voice-guidelines) no seu suplemento.

- Preocupe-se com a acessibilidade ? facilite a intera??o dos usu?rios com o seu suplemento e inclua tecnologias auxiliares, como leitores de tela.

- Desenvolva para todas as plataformas e m?todos de entrada, incluindo teclado/mouse e [toque](#optimize-for-touch). Certifique-se de que sua interface do usu?rio seja responsiva a diferentes fatores de forma.

Para modelos que aplicam os princ?pios de design que voc? pode usar e personalizar durante o desenvolvimento do suplemento, consulte [padr?es de design UX para Suplementos do Office](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code).

### <a name="optimize-for-touch"></a>Otimizar para toque

- Use a propriedade [Context.touchEnabled](https://dev.office.com/reference/add-ins/shared/office.context.touchenabled) para descobrir se o aplicativo host que executa o suplemento est? habilitado para toque.

  > [!NOTE]
  > Essa propriedade n?o tem suporte no Outlook.

- Verifique se todos os controles s?o dimensionados adequadamente para intera??o por toque. Por exemplo, se os bot?es t?m destinos de toque adequados e se as caixas de entrada t?m a dimens?o correta para que os usu?rios insiram entradas.

- N?o confie nos m?todos de entrada sem toque, como passar o cursor ou clicar com o bot?o direito do mouse.

- Verifique se o suplemento funciona nos modos retrato e paisagem. Observe que em dispositivos de toque, parte do suplemento pode ficar oculta pelo teclado virtual.

- Teste seu suplemento em um dispositivo real usando o [sideload](../testing/sideload-an-office-add-in-on-ipad-and-mac.md).

> [!NOTE]
> Se voc? est? usando o [Office UI Fabric](https://github.com/OfficeDev/Office-UI-Fabric) nos seus elementos de design, muitos desses elementos j? foram tratados.


## <a name="optimize-and-monitor-add-in-performance"></a>Otimizar e monitorar o desempenho do suplemento

- Crie a percep??o de respostas r?pidas da interface do usu?rio. Seu suplemento dever? ser carregado em 500 ms ou menos.

- Certifique-se de que todas as intera??es do usu?rio respondam em menos de um segundo.

-  Forne?a indicadores de carregamento para opera??es com longa execu??o.

- Use uma CDN para hospedar imagens, recursos e bibliotecas comuns. Carregue o m?ximo poss?vel de um s? lugar.

- Siga as pr?ticas da Web padr?o para otimizar a p?gina. Use apenas vers?es reduzidas das bibliotecas na produ??o. Carregue somente os recursos que voc? precisar e otimize como os recursos s?o carregados.

- Se o tempo de execu??o das opera??es demorar, forne?a feedback aos usu?rios. Observe os limites relacionados na tabela a seguir. Saiba mais em [Limites de recurso e otimiza??o de desempenho para Suplementos do Office](../concepts/resource-limits-and-performance-optimization.md).

  |**Classe de intera??o**|**Destino**|**Limite superior**|**Percep??o humana**|
  |:-----|:-----|:-----|:-----|
  |Instant?neo|<=50 ms|100 ms|Nenhum atraso consider?vel.|
  |R?pida|50 ? 100 ms.|200 ms|Atraso m?nimo consider?vel. N?o s?o necess?rios coment?rios.|
  |T?pico|100 ? 300 ms|500 ms|R?pido, mas n?o o suficiente para ser descrito como r?pido. N?o s?o necess?rios coment?rios.|
  |Din?mico|300 ? 500 ms.|1 segundo|N?o muito r?pido, embora pare?a ser din?mico. N?o s?o necess?rios coment?rios.|
  |Cont?nuo|>500 ms|5 segundos|Tempo de espera m?dio, j? n?o parece ser din?mico. Podem ser necess?rios coment?rios.|
  |Cativo|>500 ms|10 segundos|Longo, mas n?o o suficiente para fazer executar outra a??o. Podem ser necess?rios coment?rios.|
  |Estendida|>500 ms|>10 segundos|Longo o suficiente para realizar outra a??o durante o tempo de espera. Podem ser necess?rios coment?rios.|
  |Execu??o longa|>5 ms|>1 minuto|Os usu?rios certamente far?o algo mais.|

- Monitore a integridade do servi?o e use a telemetria para monitorar o sucesso do usu?rio.


## <a name="market-your-add-in"></a>Comercializar seu suplemento

- Publique seu suplemento no [AppSource](https://docs.microsoft.com/en-us/office/dev/store/submit-to-the-office-store) e [promova-o](https://docs.microsoft.com/en-us/office/dev/store/promote-your-office-store-solution) pelo seu site. Crie uma [listagem eficaz do AppSource](https://docs.microsoft.com/en-us/office/dev/store/create-effective-office-store-listings).

- Use t?tulos sucintos e descritivos para o suplemento. Inclua no m?ximo 128 caracteres.

- Escreva descri??es curtas e atraentes para o seu suplemento. Responda a pergunta "Qual problema este suplemento resolve?".

- Transmita a proposta de valor do seu suplemento em seu t?tulo e descri??o. N?o confie apenas em sua marca.

- Crie um site para ajudar os usu?rios a encontrar e utilizar seu suplemento.

## <a name="see-also"></a>Veja tamb?m

- [Vis?o geral da plataforma Suplementos do Office](../overview/office-add-ins.md)
