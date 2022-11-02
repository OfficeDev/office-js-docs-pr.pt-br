---
title: Teste de usabilidade de Suplementos do Office
description: Saiba como testar seu design de suplemento com usuários reais.
ms.date: 07/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: 49a2af983615779160886961e8269e4588d0fc9e
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810278"
---
# <a name="usability-testing-for-office-add-ins"></a>Teste de usabilidade de Suplementos do Office

A great add-in design takes user behaviors into account. Because your own preconceptions influence your design decisions, it’s important to test designs with real users to make sure that your add-ins work well for your customers.

Você pode executar testes de usabilidade de diferentes maneiras. Para muitos desenvolvedores de suplementos, os estudos de usabilidade remotos e nãomoderados são os mais demorados e econômicos. Vários serviços de teste populares facilitam isso; a seguir estão alguns exemplos.

- [UserTesting.com](https://www.UserTesting.com)
- [Optimalworkshop.com](https://www.Optimalworkshop.com)
- [Userzoom.com](https://www.Userzoom.com)

Esses serviços de teste o ajudam a simplificar a criação do plano de teste e remover a necessidade de buscar participantes ou moderar os testes.

You need only five participants to uncover most usability issues in your design. Incorporate small tests regularly throughout your development cycle to ensure that your product is user-centered.

> [!NOTE]
> We recommend that you test the usability of your add-in across multiple platforms. To [publish your add-in to AppSource](/office/dev/store/submit-to-appsource-via-partner-center), it must work on all [platforms that support the methods that you define](/javascript/api/requirement-sets).

## <a name="1-sign-up-for-a-testing-service"></a>1. Inscreva-se para um serviço de teste

Saiba mais em [Seleção de uma ferramenta online para o teste de usuário remoto não moderado](https://www.nngroup.com/articles/unmoderated-user-testing-tools/).

## <a name="2-develop-your-research-questions"></a>2. Desenvolva as perguntas da sua pesquisa

Research questions define the objectives of your research and guide your test plan. Your questions will help you identify participants to recruit and the tasks they will perform. Make your research questions as specific as you can. You can also seek to answer broader questions.

A seguir estão alguns exemplos de perguntas de pesquisa.

**Específicas**

- Os usuários percebem o link "avaliação gratuita" na página inicial?
- Quando os usuários inserem conteúdo do suplemento em seu documento eles entendem onde é inserido no documento?

**Amplas**

- Quais são os pontos mais problemáticos para usuário em nosso suplemento?
- Os usuários entendem o significado dos ícones na barra de comandos, antes de clicar neles?
- Os usuários localizam o menu configurações com facilidade?

É importante obter dados de toda a jornada do usuário – da descoberta do suplemento à instalação e utilização dele. Considere perguntas de pesquisa que abordam os seguintes aspectos da experiência do usuário do suplemento.

- Localização do suplemento na Loja
- Escolha da instalação do suplemento
- Experiência de primeira execução
- Comandos da faixa de opções
- Interface do Usuário do Suplemento
- Como o suplemento interage com o espaço do documento do aplicativo do Office
- O nível de controle que o usuário tem nos fluxos de inserção de conteúdo

Saiba mais em [Coleta de respostas concretas versus dados subjetivos](https://help.usertesting.com/hc/articles/115003378572-Writing-effective-questions).

## <a name="3-identify-participants-to-target"></a>3. Identifique os participantes que serão o alvo

Remote testing services can give you control over many characteristics of your test participants. Think carefully about what kinds of users you want to target. In your early stages of data collection, it might be better to recruit a wide variety of participants to identify more obvious usability issues. Later, you might choose to target groups like advanced Office users, particular occupations, or specific age ranges.

## <a name="4-create-the-participant-screener"></a>4. Crie o verificador de participantes

The screener is the set of questions and requirements you will present to prospective test participants to screen them for your test. Keep in mind that participants for services like UserTesting.com have a financial interest in qualifying for your test. It's a good idea to include trick questions in your screener if you want to  exclude certain users from the test. 

Por exemplo, se deseja encontrar participantes que estão familiarizados com o GitHub, para filtrar os usuários que possam se mostrar incorretamente, inclua respostas falsas na lista de possíveis respostas.

**Com quais dos seguintes repositórios de código fonte você tem familiaridade?**  
 a. SourceShelf  [*Reject*]  
 b. CodeContainer  [*Reject*]  
 c. GitHub  [*Must select*]  
 d. BitBucket  [*May select*]  
 e. CloudForge  [*May select*]  

Se estiver planejando testar uma compilação em funcionamento do suplemento, as perguntas a seguir podem verificar os usuários que conseguirão fazer isso.

**Este teste requer a versão mais recente do Microsoft PowerPoint. Você tem a versão mais recente do PowerPoint?**  
 a. Sim [*Deve selecionar*]  
 b. No [*Reject*]  
 c. I don’t know [*Reject*]  

**Este teste requer a instalação de um suplemento gratuito para o PowerPoint e a criação de uma conta gratuita para usá-lo. Deseja instalar um suplemento e criar uma conta gratuita?**  
 a. Sim [*Deve selecionar*]  
 b. No [*Reject*]  

Saiba mais em [Práticas recomendadas do verificador de perguntas](https://help.usertesting.com/hc/articles/115003370731-Screener-question-best-practices).

## <a name="5-create-tasks-and-questions-for-participants"></a>5. Crie tarefas e perguntas para os participantes

Try to prioritize what you want tested so that you can limit the number of tasks and questions for the participant. Some services pay participants only for a set amount of time, so you want to make sure not to go over.

Try to observe participant behaviors instead of asking about them, whenever possible. If you need to ask about behaviors, ask about what participants have done in the past, rather than what they would expect to do in a situation. This tends to give more reliable results.

The main challenge in unmoderated testing is making sure your participants understand your tasks and scenarios. Your directions should be *clear and concise*. Inevitably, if there is potential for confusion, someone will be confused.

Don't assume that your user will be on the screen they’re supposed to be on at any given point during the test. Consider telling them what screen they need to be on to start the next task.

Saiba mais em [Como escrever tarefas excelentes](https://help.usertesting.com/hc/articles/115003371651-Writing-great-tasks).

## <a name="6-create-a-prototype-to-match-the-tasks-and-questions"></a>6. Crie um protótipo para corresponder às tarefas e perguntas

Você pode testar o suplemento em funcionamento ou testar um protótipo. Observe que se você desejar testar o suplemento em funcionamento, será necessário buscar participantes que tenham a versão mais recente do Office, que estejam dispostos a instalar o suplemento e a criar uma conta (a menos que você tenha as credenciais de logon para fornecer). Depois será preciso garantir que o suplemento foi instalado com êxito.

On average, it takes about 5 minutes to walk users through how to install an add-in. The following is an example of clear, concise installation steps. Adjust the steps based on the specifics of your test.

**Instale o suplemento (insira o nome do suplemento aqui) para o PowerPoint, usando as instruções a seguir.**

1. Abra o Microsoft PowerPoint.
1. Selecione **Apresentação em Branco.**
1. Vá para **Inserir** > **Meus Suplementos**.
1. Na janela pop-up, escolha **Armazenar**.
1. Digite (Nome do suplemento) na caixa de pesquisa.
1. Escolha (Nome do suplemento).
1. Tire um momento para observar a página da Loja de forma a se familiarizar com o suplemento.
1. Escolha **Adicionar** para instalar o suplemento.

You can test a prototype at any level of interaction and visual fidelity. For more complex linking and interactivity, consider a prototyping tool like [InVision](https://www.invisionapp.com). If you just want to test static screens, you can host images online and send participants the corresponding URL, or give them a link to an online PowerPoint presentation. 

## <a name="7-run-a-pilot-test"></a>7. Execute um teste piloto

It can be tricky to get the prototype and your task/question list right. Users might be confused by tasks, or might get lost in your prototype. You should run a pilot test with 1-3 users to work out the inevitable issues with the test format. This will help to ensure that your questions are clear, that the prototype is set up correctly, and that you’re capturing the type of data you’re looking for.

## <a name="8-run-the-test"></a>8. Execute o teste

After you order your test, you will get email notifications when participants complete it. Unless you’ve targeted a specific group of participants, the tests are usually completed within a few hours.

## <a name="9-analyze-results"></a>9. Analise os resultados

This is the part where you try to make sense of the data you’ve collected. While watching the test videos, record notes about problems and successes the user has. Avoid trying to interpret the meaning of the data until you have viewed all the results.

A single participant having a usability issue is not enough to warrant making a change to the design. Two or more participants encountering the same issue suggests that other users in the general population will also encounter that issue.

In general, be careful about how you use your data to draw conclusions. Don’t fall into the trap of trying to make the data fit a certain narrative; be honest about what the data actually proves, disproves, or simply fails to provide any insight about. Keep an open mind; user behavior frequently defies designer’s expectations.

## <a name="see-also"></a>Confira também

- [Como conduzir testes de usabilidade](https://whatpixel.com/howto-conduct-usability-testing/)  
- [Práticas recomendadas para UserTesting](https://help.usertesting.com/hc/articles/115003370231-Best-practices-for-UserTesting)  
- [Minimizar desvio](https://downloads.usertesting.com/white_papers/TipSheet_MinimizingBias.pdf)  
