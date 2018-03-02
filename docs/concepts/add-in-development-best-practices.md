---
title: Práticas recomendadas para o desenvolvimento de suplementos do Office
description: ''
ms.date: 01/23/2018
---



# <a name="best-practices-for-developing-office-add-ins"></a>Práticas recomendadas para o desenvolvimento de suplementos do Office

Os suplementos eficazes oferecem uma funcionalidade exclusiva e fascinante que estende os aplicativos do Office de uma maneira visualmente atraente. Para criar um excelente suplemento, ofereça uma primeira experiência envolvente para seus usuários, desenvolva uma experiência de interface de usuário de alto nível e otimize o desempenho do seu suplemento. Aplique as práticas recomendadas descritas neste artigo para criar suplementos que ajudem os usuários a concluir suas tarefas de forma rápida e eficiente.

> [!NOTE]
> Caso pretenda [publicar](../publish/publish.md) o suplemento na experiência do Office depois de criá-lo, verifique se você está em conformidade com as [Políticas de validação do AppSource](https://docs.microsoft.com/pt-br/office/dev/store/validation-policies). Por exemplo, para passar na validação, seu suplemento deve funcionar em todas as plataformas com suporte aos métodos que você definir (para mais informações, confira a [seção 4.12](https://docs.microsoft.com/pt-br/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) e a [Página de hospedagem e disponibilidade de suplementos do Office](../overview/office-add-in-availability.md)). 

## <a name="provide-clear-value"></a>Fornecer um valor claro

- Crie suplementos que ajudem os usuários a concluir tarefas de forma rápida e eficiente. Concentre-se nos cenários que fazem sentido para aplicativos do Office. Por exemplo:
 - Torne as principais tarefas de criação mais rápidas e fáceis, com menos interrupções.
 - Habilite novos cenários no Office.
 - Incorpore serviços complementares nos hosts do Office.
 - Melhore a experiência do Office para aumentar a produtividade.
- Certifique-se de que o valor do seu suplemento seja claro para os usuários desde o princípio, [criando uma experiência envolvente na primeira execução](#create-an-engaging-first-run-experience).
- Crie uma [listagem eficaz do AppSource](https://docs.microsoft.com/pt-br/office/dev/store/create-effective-office-store-listings). Deixe claro quais são os benefícios do seu suplemento no título e na descrição. Não dependa da sua marca para dizer o que seu suplemento faz.


## <a name="create-an-engaging-first-run-experience"></a>Criar uma experiência envolvente na primeira execução

- Envolva os novos usuários com uma primeira experiência altamente útil e intuitiva. Observe que, mesmo depois de baixar o suplemento da loja, os usuários ainda estão decidindo se vão utilizá-lo.

- Deixe claro quais são as etapas que usuário terá que seguir para se envolver com seu suplemento. Use vídeos, diagramas, painéis de paginação ou outros recursos para atrair usuários.

- Reforce a proposta de valor do seu suplemento no início, em vez de apenas pedir que seus usuários entrem.

- Forneça uma interface do usuário informativa e torne sua interface do usuário pessoal.

   ![Uma captura de tela que mostra um painel de tarefas de suplemento com etapas de introdução ao lado de um suplemento sem etapas de introdução](../images/contoso-part-catalog-do-dont.png)

- Se seu suplemento de conteúdo estiver vinculado a dados no documento do usuário, inclua exemplos de dados ou um modelo para mostrar aos usuários o formato de dados a ser usado.

   ![Uma captura de tela que mostra um suplemento de conteúdo com dados ao lado de um suplemento de conteúdo sem dados](../images/add-in-title.png)

- Ofereça [avaliações gratuitas](https://docs.microsoft.com/pt-br/office/dev/store/decide-on-a-pricing-model#office-store-pricing-options). Caso o suplemento exija uma assinatura, disponibilize algumas funcionalidades sem a necessidade da assinatura.

- Simplifique o processo de inscrição. Preencha automaticamente as informações (email, nome de exibição) e ignore as verificações de email.

- Evite os pop-ups. Se você tiver de usá-los, oriente o usuário sobre como habilitar o seu pop-up.

- Use a [autenticação de logon único (SSO)](https://docs.microsoft.com/pt-br/outlook/add-ins/authenticate-a-user-with-an-identity-token).

Para modelos que ilustram padrões que podem ser aplicados enquanto você desenvolve sua experiência na primeira execução, consulte [padrões de design UX para Suplementos do Office](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code).

## <a name="use-add-in-commands"></a>Usar comandos de suplemento

- Fornece ao suplemento pontos de entrada relevantes da interface do usuário usando os comandos do suplemento. Confira mais detalhes, inclusive as práticas recomendadas de design, nos [comandos de suplemento](../design/add-in-commands.md).

## <a name="apply-ux-design-principles"></a>Aplicar os princípios de design de UX

- Assegure-se de que a aparência e a funcionalidade de seus suplementos complementam a experiência do Office. Use o [Office UI Fabric](https://dev.office.com/fabric).

- Favoreça o conteúdo através do Chrome. Evite elementos de interface do usuário supérfluos que não agregam valor à experiência do usuário.

- Mantenha os usuários no controle. Verifique se os usuários compreenderam as decisões importantes e podem reverter facilmente as ações realizadas pelo suplemento.

- Use uma identidade visual para inspirar confiança e orientar os usuários. Não use o recurso de identidade visual para sobrecarregar ou enviar anúncios aos usuários.

- Evite a necessidade de rolagem. Otimize para a resolução 1366 x 768.

- Não inclua imagens não licenciadas.

- Use uma [linguagem clara e simples](../design/add-in-design-guidelines.md#voice-guidelines) no seu suplemento.

- Preocupe-se com a [acessibilidade](../design/ui-elements/accessibility-guidelines.md) – facilite a interação dos usuários com o seu suplemento e inclua tecnologias auxiliares, como leitores de tela.

- Desenvolva para todas as plataformas e métodos de entrada, incluindo teclado/mouse e [toque](#optimize-for-touch). Certifique-se de que sua interface do usuário seja responsiva a diferentes fatores de forma.

Para modelos que aplicam os princípios de design que você pode usar e personalizar durante o desenvolvimento do suplemento, consulte [padrões de design UX para Suplementos do Office](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code).

### <a name="optimize-for-touch"></a>Otimizar para toque

- Use a propriedade [Context.touchEnabled](https://dev.office.com/reference/add-ins/shared/office.context.touchenabled) para descobrir se o aplicativo host que executa o suplemento está habilitado para toque.

  > [!NOTE]
  > Essa propriedade não tem suporte no Outlook.

- Verifique se todos os controles são dimensionados adequadamente para interação por toque. Por exemplo, se os botões têm destinos de toque adequados e se as caixas de entrada têm a dimensão correta para que os usuários insiram entradas.

- Não confie nos métodos de entrada sem toque, como passar o cursor ou clicar com o botão direito do mouse.

- Verifique se o suplemento funciona nos modos retrato e paisagem. Observe que em dispositivos de toque, parte do suplemento pode ficar oculta pelo teclado virtual.

- Teste seu suplemento em um dispositivo real usando o [sideload](../testing/sideload-an-office-add-in-on-ipad-and-mac.md).

> [!NOTE]
> Se você está usando o [Office UI Fabric](https://github.com/OfficeDev/Office-UI-Fabric) nos seus elementos de design, muitos desses elementos já foram tratados.


## <a name="optimize-and-monitor-add-in-performance"></a>Otimizar e monitorar o desempenho do suplemento

- Crie a percepção de respostas rápidas da interface do usuário. Seu suplemento deverá ser carregado em 500 ms ou menos.

- Certifique-se de que todas as interações do usuário respondam em menos de um segundo.

-  Forneça indicadores de carregamento para operações com longa execução.

- Use uma CDN para hospedar imagens, recursos e bibliotecas comuns. Carregue o máximo possível de um só lugar.

- Siga as práticas da Web padrão para otimizar a página. Use apenas versões reduzidas das bibliotecas na produção. Carregue somente os recursos que você precisar e otimize como os recursos são carregados.

- Se o tempo de execução das operações demorar, forneça feedback aos usuários. Observe os limites relacionados na tabela a seguir. Saiba mais em [Limites de recurso e otimização de desempenho para Suplementos do Office](../concepts/resource-limits-and-performance-optimization.md).

  |**Classe de interação**|**Destino**|**Limite superior**|**Percepção humana**|
  |:-----|:-----|:-----|:-----|
  |Instantâneo|<=50 ms|100 ms|Nenhum atraso considerável.|
  |Rápida|50 – 100 ms.|200 ms|Atraso mínimo considerável. Não são necessários comentários.|
  |Típico|100 – 300 ms|500 ms|Rápido, mas não o suficiente para ser descrito como rápido. Não são necessários comentários.|
  |Dinâmico|300 – 500 ms.|1 segundo|Não muito rápido, embora pareça ser dinâmico. Não são necessários comentários.|
  |Contínuo|>500 ms|5 segundos|Tempo de espera médio, já não parece ser dinâmico. Podem ser necessários comentários.|
  |Cativo|>500 ms|10 segundos|Longo, mas não o suficiente para fazer executar outra ação. Podem ser necessários comentários.|
  |Estendida|>500 ms|>10 segundos|Longo o suficiente para realizar outra ação durante o tempo de espera. Podem ser necessários comentários.|
  |Execução longa|>5 ms|>1 minuto|Os usuários certamente farão algo mais.|

- Monitore a integridade do serviço e use a telemetria para monitorar o sucesso do usuário.


## <a name="market-your-add-in"></a>Comercializar seu suplemento

- Publique seu suplemento no [AppSource](https://docs.microsoft.com/pt-br/office/dev/store/submit-to-the-office-store) e [promova-o](https://docs.microsoft.com/pt-br/office/dev/store/promote-your-office-store-solution) pelo seu site. Crie uma [listagem eficaz do AppSource](https://docs.microsoft.com/pt-br/office/dev/store/create-effective-office-store-listings).

- Use títulos sucintos e descritivos para o suplemento. Inclua no máximo 128 caracteres.

- Escreva descrições curtas e atraentes para o seu suplemento. Responda a pergunta "Qual problema este suplemento resolve?".

- Transmita a proposta de valor do seu suplemento em seu título e descrição. Não confie apenas em sua marca.

- Crie um site para ajudar os usuários a encontrar e utilizar seu suplemento.

## <a name="see-also"></a>Veja também

- [Visão geral da plataforma Suplementos do Office](../overview/office-add-ins.md)
