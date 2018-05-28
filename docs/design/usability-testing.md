---
title: Teste de usabilidade de Suplementos do Office
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 410b8d7ede22cf222ee2df794e438c7f5f8881dd
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="usability-testing-for-office-add-ins"></a>Teste de usabilidade de Suplementos do Office

Um excelente design de suplemento considera os comportamentos do usu?rio. Como seus pr?prios conceitos pr?vios influenciam suas decis?es de design, ? importante testar designs com usu?rios reais para garantir que seus suplementos funcionem bem para seus clientes. 

? poss?vel executar testes de usabilidade de maneiras diferentes. Para muitos desenvolvedores de suplementos, estudos de usabilidade remota n?o moderada s?o os que economizam mais tempo e dinheiro. V?rios servi?os de testes populares facilitam isso. Veja alguns exemplos: 

 - [UserTesting.com](https://www.UserTesting.com)
 - [Optimalworkshop.com](https://www.Optimalworkshop.com)
 - [Userzoom.com](https://www.Userzoom.com)

Esses servi?os de teste o ajudam a simplificar a cria??o do plano de teste e remover a necessidade de buscar participantes ou moderar os testes. 

Voc? precisa de apenas cinco participantes para descobrir a maioria dos problemas de usabilidade no seu design. Incorpore testes pequenos regularmente durante o ciclo de desenvolvimento para garantir que seu produto seja centralizado no usu?rio.

> [!NOTE]
> Recomendamos que voc? teste a usabilidade do seu suplemento em v?rias plataformas. Para [publicar](https://docs.microsoft.com/en-us/office/dev/store/submit-to-the-office-store) seu suplemento no AppSource, ele deve funcionar em todas as [plataformas compat?veis com os m?todos que voc? definir](../overview/office-add-in-availability.md).

## <a name="1---sign-up-for-a-testing-service"></a>1.   Inscreva-se em um servi?o de teste

Saiba mais em [Sele??o de uma ferramenta online para o teste de usu?rio remoto n?o moderado.](https://www.nngroup.com/articles/unmoderated-user-testing-tools/)

## <a name="2-develop-your-research-questions"></a>2. Desenvolva as perguntas da sua pesquisa
 
As perguntas da pesquisa definem os objetivos de sua pesquisa e guiam seu plano de teste. Suas perguntas o ajudar?o a identificar os participantes para recrutar e as tarefas que eles executar?o. Certifique-se de que suas perguntas de pesquisa sejam o mais espec?ficas poss?vel. Voc? tamb?m pode procurar responder perguntas mais amplas.
 
A seguir, alguns exemplos de perguntas de pesquisa:
  
**Espec?ficas**  

 - Os usu?rios percebem o link "avalia??o gratuita" na p?gina inicial?
 - Quando os usu?rios inserem conte?do do suplemento em seu documento eles entendem onde ? inserido no documento?

**Amplas**  

 - Quais s?o os pontos mais problem?ticos para usu?rio em nosso suplemento?
 - Os usu?rios entendem o significado dos ?cones na barra de comandos, antes de clicar neles?
 - Os usu?rios localizam o menu configura??es com facilidade?

? importante obter dados de toda a jornada do usu?rio ? da descoberta do suplemento ? instala??o e utiliza??o dele. Considere perguntas de pesquisa que abordem os seguintes aspectos da experi?ncia do usu?rio no suplemento:
 
 - Localiza??o do suplemento na Loja
 - Escolha da instala??o do suplemento
 - Experi?ncia de primeira execu??o
 - Comandos da faixa de op??es
 - Interface do Usu?rio do Suplemento
 - Como o suplemento interage com o espa?o do documento do aplicativo do Office
 - Qual o n?vel de controle que o usu?rio tem nos fluxos de inser??o de conte?do

Para saber mais, veja [Escrever perguntas eficazes.](http://help.usertesting.com/customer/en/portal/articles/2077663-writing-effective-questions)
 
## <a name="3-identify-participants-to-target"></a>3. Identifique os participantes que ser?o o alvo
 
O teste remoto de servi?os pode oferecer a voc? o controle de v?rias caracter?sticas dos participantes do teste. Pense cuidadosamente sobre que tipos de usu?rios voc? deseja buscar. Nos seus est?gios iniciais de coleta de dados, talvez seja melhor recrutar uma ampla variedade de participantes para identificar problemas de usabilidade mais ?bvios. Posteriormente, voc? pode optar por grupos segmentados como usu?rios avan?ados do Office, ocupa??es espec?ficas ou faixas et?rias espec?ficas.
 
## <a name="4-create-the-participant-screener"></a>4. Crie o verificador de participantes
 
O verificador ? o conjunto de perguntas e requisitos que voc? apresentar? aos participantes do teste em potencial para verific?-los para o teste. Tenha em mente que os participantes de servi?os como UserTesting.com t?m interesse financeiro em se qualificar para seu teste. ? uma boa ideia incluir perguntas dif?ceis em sua verifica??o se desejar excluir determinados usu?rios do teste. 
 
Por exemplo, se deseja encontrar participantes que est?o familiarizados com o GitHub, para filtrar os usu?rios que possam se mostrar incorretamente, inclua respostas falsas na lista de poss?veis respostas.

**Com quais dos seguintes reposit?rios de c?digo fonte voc? tem familiaridade?**  
 a. SourceShelf [*Rejeitar*]  
 b. CodeContainer [*Rejeitar*]  
 c. GitHub [*Deve selecionar*]  
 d. BitBucket [*Pode selecionar*]  
 e. CloudForge [*Pode selecionar*]  

Se estiver planejando testar uma compila??o em funcionamento do suplemento, as perguntas a seguir podem verificar os usu?rios que conseguir?o fazer isso. 

**Este teste requer que voc? tenha o Microsoft PowerPoint 2016. Voc? tem o PowerPoint 2016?**  
 a. Sim [*Deve selecionar*]  
 b. N?o [*Rejeitar*]  
 c. N?o sei [*Rejeitar*]  

**Este teste requer que voc? instale um suplemento gratuito para PowerPoint 2016 e crie uma conta gratuita para us?-lo. Voc? est? disposto a instalar um suplemento e criar uma conta gratuita?**  
 a. Sim [*Deve selecionar*]  
 b. N?o [*Rejeitar*]  

Para saber mais, veja [Pr?ticas recomendadas do verificador de perguntas.](http://help.usertesting.com/customer/en/portal/articles/2077835-screener-question-best-practices)
 
## <a name="5-create-tasks-and-questions-for-participants"></a>5. Crie tarefas e perguntas para os participantes
 
Tente priorizar o que voc? quer testar para que seja poss?vel limitar o n?mero de tarefas e perguntas do participante. Alguns servi?os pagam os participantes apenas para um determinado per?odo para que voc? certifique-se de n?o exced?-lo.

Tente observar como os participantes se comportam em vez de perguntar sobre eles sempre que poss?vel. Se voc? precisar perguntar sobre comportamentos, pergunte o que os participantes fizeram no passado, em vez do que o que eles esperariam fazer em uma situa??o. Isso tende a fornecer resultados mais confi?veis.
 
O principal desafio no teste n?o moderado ? garantir que seus participantes entendam suas tarefas e cen?rios. Suas orienta??es devem ser *claras e concisas*. Inevitavelmente, se houver potencial para confus?o, algu?m ficar? confuso. 

N?o pense que o usu?rio estar? na tela que deve estar em um determinado momento durante o teste. Considere informar a tela em que eles precisam estar para iniciar a pr?xima tarefa. 

Saiba mais em [Como escrever tarefas excelentes.](http://help.usertesting.com/customer/en/portal/articles/2077824-writing-great-tasks)

## <a name="6-create-a-prototype-to-match-the-tasks-and-questions"></a>6. Crie um prot?tipo para corresponder ?s tarefas e perguntas
 
Voc? tamb?m pode testar o suplemento em funcionamento ou pode testar um prot?tipo. Tenha em mente que, se voc? deseja testar o suplemento em funcionamento, ? necess?rio buscar participantes que tenham o Office 2016, estejam dispostos a instalar o suplemento e criar uma conta (a menos que voc? tenha as credenciais de logon para fornecer). Voc? precisar? certificar-se de que o suplemento ser? instalado com ?xito. 

Em m?dia, s?o necess?rios cerca de cinco minutos para orientar os usu?rios sobre como instalar um suplemento. A seguir, um exemplo de etapas de instala??o claras e concisas. Ajuste as etapas com base nas condi??es espec?ficas do teste.

**Instale o suplemento (insira o nome do suplemento aqui) para o PowerPoint 2016 usando as seguintes instru??es:** 

1. Abra o Microsoft PowerPoint 2016.
2. Selecione **Apresenta??o em Branco.**
3. V? para **Inserir > Meus Suplementos.**
5. Na janela pop-up, escolha **Loja.**
6. Digite (Nome do suplemento) na caixa de pesquisa.
7. Escolha (Nome do suplemento).
8. Tire um momento para observar a p?gina da Loja de forma a se familiarizar com o suplemento.
9. Escolha **Adicionar** para instalar o suplemento.

Voc? pode testar um prot?tipo em qualquer n?vel de intera??o e fidelidade visual. Para vincula??o e interatividade mais complexas, considere uma ferramenta de cria??o de prot?tipo como a [InVision](https://www.invisionapp.com). Se voc? deseja testar telas est?ticas, ? poss?vel hospedar imagens online e enviar a URL correspondente para os participantes ou fornecer um link para uma apresenta??o online do PowerPoint. 

## <a name="7-run-a-pilot-test"></a>7. Execute um teste piloto

Pode ser dif?cil acertar no prot?tipo e na lista de tarefas/perguntas. Os usu?rios podem ficar confusos com as tarefas ou podem se perder em seu prot?tipo. Voc? deve fazer um teste piloto 1 a 3 usu?rios para trabalhar corrigir os inevit?veis problemas com o formato do teste. Isso ajudar? a garantir que suas perguntas sejam claras, que o prot?tipo esteja configurado corretamente e que voc? esteja capturando o tipo de dados que est? procurando.

## <a name="8-run-the-test"></a>8. Execute o teste

Depois que voc? solicitar o teste, receber? notifica??es por email quando os participantes o conclu?rem. A menos que tenha direcionado para um grupo espec?fico de participantes, os testes normalmente s?o conclu?dos dentro de algumas horas.

## <a name="9-analyze-results"></a>9. Analise os resultados

Essa ? a parte em que voc? tenta fazer com que os dados coletados fa?am sentido. Ao assistir os v?deos de teste, anote os problemas e os ?xitos do usu?rio. Evite tentar interpretar o significado dos dados at? que tenha exibido todos os resultados. 

Um ?nico participante com um problema de usabilidade n?o ? suficiente para gerar uma altera??o no design. Dois ou mais participantes que encontram o mesmo problema sugere que outros usu?rios no geral tamb?m encontrar?o esse problema.

Em geral, tome cuidado com como voc? usa seus dados para tirar conclus?es. N?o caia na armadilha de tentar fazer com que os dados se ajustem a uma determinada narrativa. Seja honesto sobre o que os dados realmente comprovam, refutam ou apenas falham em oferecer informa??es. Mantenha a mente aberta. O comportamento do usu?rio com frequ?ncia desafia as expectativas do designer.
 

## <a name="see-also"></a>Veja tamb?m
 
 - [Como conduzir testes de usabilidade](http://whatpixel.com/howto-conduct-usability-testing/)  
 - [Pr?ticas recomendadas](http://help.usertesting.com/customer/en/portal/articles/1680726-best-practices)  
 - [Minimizar desvio](http://downloads.usertesting.com/white_papers/TipSheet_MinimizingBias.pdf)  
