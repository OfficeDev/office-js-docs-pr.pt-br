---
ms.date: 02/20/2020
title: Configure o suplemento do Excel para compartilhar o tempo de execução do navegador (visualização)
ms.prod: excel
description: Configure o suplemento do Excel para compartilhar o tempo de execução do navegador e executar a faixa de opções, o painel de tarefas e o código de função personalizado no mesmo tempo de execução.
localization_priority: Priority
ms.openlocfilehash: 7945bd8fdb29a9d6d44d7d29676410a54bacf83f
ms.sourcegitcommit: dd6d00202f6466c27418247dad7bd136555a6036
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/26/2020
ms.locfileid: "42284099"
---
# <a name="configure-your-excel-add-in-to-use-a-shared-javascript-runtime-preview"></a>Configure seu suplemento do Excel para usar um tempo de execução JavaScript compartilhado (visualização)

[!include[Running custom functions in a shared runtime note](../includes/excel-shared-runtime-preview-note.md)]

Ao executar o Excel no Windows ou Mac, o suplemento executará o código para botões da faixa de opções, funções personalizadas e o painel de tarefas em ambientes de tempo de execução JavaScript separados. Isso cria limitações, como não poder compartilhar facilmente dados globais e não poder acessar todas as funcionalidades do CORS a partir de uma função customizada.

No entanto, você pode configurar o suplemento do Excel para compartilhar código em um tempo de execução JavaScript compartilhado. Isso permite uma melhor coordenação entre seu suplemento e acesso ao DOM e CORS de todas as partes do seu suplemento. Também permite executar o código quando o documento é aberto ou executar o código enquanto o painel de tarefas está fechado. Para configurar seu suplemento para usar um tempo de execução compartilhado, siga as instruções neste artigo.

## <a name="create-the-add-in-project"></a>Criar o projeto do suplemento

Se você estiver iniciando um novo projeto, siga estas etapas para usar o gerador Yeoman para criar um projeto de suplemento do Excel. Execute o comando a seguir e responda às solicitações com as seguintes respostas:

```command line
yo office
```

- Escolha um tipo de projeto: **Projeto de suplemento de funções personalizadas do Excel**
- Escolha um tipo de script: **JavaScript**
- Qual será o nome do seu suplemento? **Meu suplemento do Office**

![Captura de tela das solicitações de resposta do seu Office para criar o projeto de suplemento.](../images/yo-office-excel-project.png)

Depois que você concluir o assistente, o gerador criará o projeto e instalará os componentes Node de suporte.

## <a name="configure-the-manifest"></a>Configurar o manifesto

Siga estas etapas para um projeto novo ou existente para configurá-lo para usar um tempo de execução compartilhado.

1. Inicie o código do Visual Studio e abra o projeto **Meu suplemento do Office**.
2. Abra o arquivo **manifest.xml**.
3. Localize a seção `<VersionOverrides>` e adicione a seguinte seção `<Runtimes>`. O tempo de vida precisa ser **longo** para que as funções personalizadas ainda possam funcionar, mesmo quando o painel de tarefas estiver fechado. O resid é `ContosoAddin.Url`, que faz referência a uma sequência na seção de recursos posteriormente. Você pode usar qualquer valor de resid que desejar, mas deve corresponder ao resid dos outros elementos nos elementos do suplemento.

   ```xml
   <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
     <Hosts>
       <Host xsi:type="Workbook">
       <Runtimes>
         <Runtime resid="ContosoAddin.Url" lifetime="long" />
       </Runtimes>
       <AllFormFactors>
   ```

4. No elemento `<Page>`, altere o local de origem de **Functions.Page.Url** para **ContosoAddin.Url**. Este resid corresponde ao elemento resid `<Runtime>`. Observe que, se você não tiver funções personalizadas, não terá uma entrada **Page** e poderá pular esta etapa.

   ```xml
   <AllFormFactors>
   ...
   <Page>
   <SourceLocation resid="ContosoAddin.Url"/>
   </Page>
   ...
   ```

5. Na seção `<DesktopFormFactor>`, altere o **FunctionFile** de **Commands.Url** para usar **ContosoAddin.Url**. Observe que, se você não possui comandos de ação, não terá uma entrada **FunctionFile** e poderá pular esta etapa.

   ```xml
   <DesktopFormFactor>
   <GetStarted>
   ...
   </GetStarted>
   <FunctionFile resid="ContosoAddin.Url"/>
   ```

6. Na seção `<Action>`, altere o local de origem de **Taskpane.Url** para **ContosoAddin.Url**. Observe que, se você não tiver um painel de tarefas, não terá uma ação **ShowTaskpane** e poderá pular esta etapa.

   ```xml
   <Action xsi:type="ShowTaskpane">
   <TaskpaneId>ButtonId1</TaskpaneId>
   <SourceLocation resid="ContosoAddin.Url"/>
   </Action>
   ```

7. Adicione um novo **ID de URL** para **ContosoAddin.Url** que aponte para **taskpane.html**.

   ```xml
   <bt:Urls>
   <bt:Url id="Functions.Script.Url" DefaultValue="https://localhost:3000/dist/functions.js"/>
   ...
   <bt:Url id="ContosoAddin.Url" DefaultValue="https://localhost:3000/taskpane.html"/>
   ...
   ```

8. Salve suas alterações e recompile o projeto.

   ```command line
   npm run build
   ```

## <a name="runtime-lifetime"></a>Duração do tempo de execução

Ao adicionar o elemento `Runtime`, você também especifica uma vida útil com um valor de `long` ou `short`. Defina esse valor como `long` para aproveitar os recursos, como iniciar o suplemento quando o documento for aberto, continuar executando o código após o fechamento do painel de tarefas ou usar o CORS e o DOM nas funções personalizadas.

Se você definir esse valor como `short`, o suplemento se comportará de maneira semelhante ao comportamento padrão. O suplemento será iniciado quando um dos botões da faixa de opções for pressionado, mas poderá desligar após a execução do manipulador da faixa de opções. Da mesma forma, o suplemento será iniciado quando o painel de tarefas for aberto, mas poderá ser desativado quando o painel de tarefas estiver fechado.

```xml
<Runtimes>
  <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="multiple-task-panes"></a>Vários painéis de tarefas

Não projete seu suplemento para usar vários painéis de tarefas se você planeja usar o tempo de execução compartilhado. O tempo de execução compartilhado tem suporte para o uso de apenas um único painel de tarefas. Observe que qualquer painel de tarefas sem um `<TaskpaneID>` é considerado um painel de tarefas diferente.

## <a name="next-steps"></a>Próximas etapas

Agora experimente alguns recursos do tempo de execução compartilhado, observando os seguintes artigos.

- [Chamar APIs do Excel a partir de uma função personalizada](call-excel-apis-from-custom-function.md)

## <a name="see-also"></a>Confira também

- [Visão geral: Execute seu código de suplemento em um tempo de execução JavaScript compartilhado (visualização)](custom-functions-shared-overview.md)
