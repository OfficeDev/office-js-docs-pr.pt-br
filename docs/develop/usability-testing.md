---
title: Teste de usabilidade de Suplementos do Office
description: Saiba como testar o design do seu complemento com usuários reais.
ms.date: 07/08/2021
localization_priority: Normal
ms.openlocfilehash: 4f797033343fd20d814cba1fd67ad8eb21b246b3
ms.sourcegitcommit: e570fa8925204c6ca7c8aea59fbf07f73ef1a803
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/05/2021
ms.locfileid: "53773339"
---
# <a name="usability-testing-for-office-add-ins"></a>Teste de usabilidade de Suplementos do Office

Um excelente design de suplemento considera os comportamentos do usuário. Como seus próprios conceitos prévios influenciam suas decisões de design, é importante testar designs com usuários reais para garantir que seus suplementos funcionem bem para seus clientes.

Você pode executar testes de usabilidade de maneiras diferentes. Para muitos desenvolvedores de complementos, os estudos de usabilidade remotos e nãomoderados são os que têm mais tempo e custo benefício. Vários serviços de teste populares facilitam isso; a seguir estão alguns exemplos.

- [UserTesting.com](https://www.UserTesting.com)
- [Optimalworkshop.com](https://www.Optimalworkshop.com)
- [Userzoom.com](https://www.Userzoom.com)

Esses serviços de teste o ajudam a simplificar a criação do plano de teste e remover a necessidade de buscar participantes ou moderar os testes.

Você precisa de apenas cinco participantes para descobrir a maioria dos problemas de usabilidade no seu design. Incorpore testes pequenos regularmente durante o ciclo de desenvolvimento para garantir que seu produto seja centralizado no usuário.

> [!NOTE]
> Recomendamos que você teste a usabilidade do seu suplemento em várias plataformas. Para [publicar](/office/dev/store/submit-to-appsource-via-partner-center) seu suplemento no AppSource, ele deve funcionar em todas as [plataformas compatíveis com os métodos que você definir](../overview/office-add-in-availability.md).

## <a name="1-sign-up-for-a-testing-service"></a>1. Inscreva-se em um serviço de teste

Saiba mais em [Seleção de uma ferramenta online para o teste de usuário remoto não moderado](https://www.nngroup.com/articles/unmoderated-user-testing-tools/).

## <a name="2-develop-your-research-questions"></a>2. Desenvolva as perguntas da sua pesquisa

As perguntas da pesquisa definem os objetivos de sua pesquisa e guiam seu plano de teste. Suas perguntas o ajudarão a identificar os participantes para recrutar e as tarefas que eles executarão. Certifique-se de que suas perguntas de pesquisa sejam o mais específicas possível. Você também pode procurar responder perguntas mais amplas.

A seguir estão alguns exemplos de perguntas de pesquisa.

**Específicas**

- Os usuários percebem o link "avaliação gratuita" na página inicial?
- Quando os usuários inserem conteúdo do suplemento em seu documento eles entendem onde é inserido no documento?

**Amplas**

- Quais são os pontos mais problemáticos para usuário em nosso suplemento?
- Os usuários entendem o significado dos ícones na barra de comandos, antes de clicar neles?
- Os usuários localizam o menu configurações com facilidade?

É importante obter dados de toda a jornada do usuário – da descoberta do suplemento à instalação e utilização dele. Considere perguntas de pesquisa que abordam os seguintes aspectos da experiência do usuário do complemento.

- Localização do suplemento na Loja
- Escolha da instalação do suplemento
- Experiência de primeira execução
- Comandos da faixa de opções
- Interface do Usuário do Suplemento
- Como o suplemento interage com o espaço do documento do aplicativo do Office
- O nível de controle que o usuário tem nos fluxos de inserção de conteúdo

Saiba mais em [Coleta de respostas concretas versus dados subjetivos](https://help.usertesting.com/hc/articles/115003378572-Writing-effective-questions).

## <a name="3-identify-participants-to-target"></a>3. Identifique os participantes que serão o alvo

O teste remoto de serviços pode oferecer a você o controle de várias características dos participantes do teste. Pense cuidadosamente sobre que tipos de usuários você deseja buscar. Nos seus estágios iniciais de coleta de dados, talvez seja melhor recrutar uma ampla variedade de participantes para identificar problemas de usabilidade mais óbvios. Posteriormente, você pode optar por grupos segmentados como usuários avançados do Office, ocupações específicas ou faixas etárias específicas.

## <a name="4-create-the-participant-screener"></a>4. Crie o verificador de participantes

O verificador é o conjunto de perguntas e requisitos que você apresentará aos participantes do teste em potencial para verificá-los para o teste. Tenha em mente que os participantes de serviços como UserTesting.com têm interesse financeiro em se qualificar para seu teste. É uma boa ideia incluir perguntas difíceis em sua verificação se desejar excluir determinados usuários do teste. 

Por exemplo, se deseja encontrar participantes que estão familiarizados com o GitHub, para filtrar os usuários que possam se mostrar incorretamente, inclua respostas falsas na lista de possíveis respostas.

**Com quais dos seguintes repositórios de código fonte você tem familiaridade?**  
 a. SourceShelf [*Rejeitar*]  
 b. CodeContainer [*Rejeitar*]  
 c. GitHub [*Deve selecionar*]  
 d. BitBucket [*Pode selecionar*]  
 e. CloudForge [*Pode selecionar*]  

Se estiver planejando testar uma compilação em funcionamento do suplemento, as perguntas a seguir podem verificar os usuários que conseguirão fazer isso.

**Este teste requer a versão mais recente do Microsoft PowerPoint. Você tem a versão mais recente do PowerPoint?**  
 a. Sim [*Deve selecionar*]  
 b. Não [*Rejeitar*]  
 c. Não sei [*Rejeitar*]  

**Este teste requer a instalação de um suplemento gratuito para o PowerPoint e a criação de uma conta gratuita para usá-lo. Deseja instalar um suplemento e criar uma conta gratuita?**  
 a. Sim [*Deve selecionar*]  
 b. Não [*Rejeitar*]  

Saiba mais em [Práticas recomendadas do verificador de perguntas](https://help.usertesting.com/hc/articles/115003370731-Screener-question-best-practices).

## <a name="5-create-tasks-and-questions-for-participants"></a>5. Crie tarefas e perguntas para os participantes

Tente priorizar o que você quer testar para que seja possível limitar o número de tarefas e perguntas do participante. Alguns serviços pagam os participantes apenas para um determinado período para que você certifique-se de não excedê-lo.

Tente observar como os participantes se comportam em vez de perguntar sobre eles sempre que possível. Se você precisar perguntar sobre comportamentos, pergunte o que os participantes fizeram no passado, em vez do que o que eles esperariam fazer em uma situação. Isso tende a fornecer resultados mais confiáveis.

O principal desafio no teste não moderado é garantir que seus participantes entendam suas tarefas e cenários. Suas orientações devem ser *claras e concisas*. Inevitavelmente, se houver potencial para confusão, alguém ficará confuso.

Não pense que o usuário estará na tela que deve estar em um determinado momento durante o teste. Considere informar a tela em que eles precisam estar para iniciar a próxima tarefa.

Saiba mais em [Como escrever tarefas excelentes](https://help.usertesting.com/hc/articles/115003371651-Writing-great-tasks).

## <a name="6-create-a-prototype-to-match-the-tasks-and-questions"></a>6. Crie um protótipo para corresponder às tarefas e perguntas

Você pode testar o suplemento em funcionamento ou testar um protótipo. Observe que se você desejar testar o suplemento em funcionamento, será necessário buscar participantes que tenham a versão mais recente do Office, que estejam dispostos a instalar o suplemento e a criar uma conta (a menos que você tenha as credenciais de logon para fornecer). Depois será preciso garantir que o suplemento foi instalado com êxito.

Em média, são necessários cerca de cinco minutos para orientar os usuários sobre como instalar um suplemento. A seguir, um exemplo de etapas de instalação claras e concisas. Ajuste as etapas com base nas condições específicas do teste.

**Instale o complemento (insira o nome do seu complemento aqui) para PowerPoint, usando as instruções a seguir.**

1. Abra o Microsoft PowerPoint.
1. Selecione **Apresentação em Branco.**
1. Vá para **Inserir > Meus Suplementos.**
1. Na janela pop-up, escolha **Loja.**
1. Digite (Nome do suplemento) na caixa de pesquisa.
1. Escolha (Nome do suplemento).
1. Tire um momento para observar a página da Loja de forma a se familiarizar com o suplemento.
1. Escolha **Adicionar** para instalar o suplemento.

Você pode testar um protótipo em qualquer nível de interação e fidelidade visual. Para vinculação e interatividade mais complexas, considere uma ferramenta de criação de protótipo como a [InVision](https://www.invisionapp.com). Se você deseja testar telas estáticas, é possível hospedar imagens online e enviar a URL correspondente para os participantes ou fornecer um link para uma apresentação online do PowerPoint. 

## <a name="7-run-a-pilot-test"></a>7. Execute um teste piloto

Pode ser difícil acertar no protótipo e na lista de tarefas/perguntas. Os usuários podem ficar confusos com as tarefas ou podem se perder em seu protótipo. Você deve fazer um teste piloto 1 a 3 usuários para trabalhar corrigir os inevitáveis problemas com o formato do teste. Isso ajudará a garantir que suas perguntas sejam claras, que o protótipo esteja configurado corretamente e que você esteja capturando o tipo de dados que está procurando.

## <a name="8-run-the-test"></a>8. Execute o teste

Depois que você solicitar o teste, receberá notificações por email quando os participantes o concluírem. A menos que tenha direcionado para um grupo específico de participantes, os testes normalmente são concluídos dentro de algumas horas.

## <a name="9-analyze-results"></a>9. Analise os resultados

Essa é a parte em que você tenta fazer com que os dados coletados façam sentido. Ao assistir os vídeos de teste, anote os problemas e os êxitos do usuário. Evite tentar interpretar o significado dos dados até que tenha exibido todos os resultados.

Um único participante com um problema de usabilidade não é suficiente para gerar uma alteração no design. Dois ou mais participantes que encontram o mesmo problema sugere que outros usuários no geral também encontrarão esse problema.

Em geral, tome cuidado com como você usa seus dados para tirar conclusões. Não caia na armadilha de tentar fazer com que os dados se ajustem a uma determinada narrativa. Seja honesto sobre o que os dados realmente comprovam, refutam ou apenas falham em oferecer informações. Mantenha a mente aberta. O comportamento do usuário com frequência desafia as expectativas do designer.

## <a name="see-also"></a>Confira também

- [Como conduzir testes de usabilidade](https://whatpixel.com/howto-conduct-usability-testing/)  
- [Práticas recomendadas para UserTesting](https://help.usertesting.com/hc/en-us/articles/115003370231-Best-practices-for-UserTesting)  
- [Minimizar desvio](https://downloads.usertesting.com/white_papers/TipSheet_MinimizingBias.pdf)  
