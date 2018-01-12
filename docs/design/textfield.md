# <a name="textfield-component-in-office-ui-fabric"></a>Componente TextField no Office UI Fabric

Um campo de texto permite aos usuários digitar texto. Geralmente, são usados para capturar uma única linha de texto, mas podem ser configurados para capturar várias linhas. O texto é exibido na tela em um formato simples e uniforme.
  
#### <a name="example-textfield-in-a-task-pane"></a>Exemplo: TextField em um painel de tarefas

![Imagem mostrando o TextField](../../images/overview_withApp_textField.png)

<br/>

## <a name="best-practices"></a>Práticas recomendadas

|**Faça**|**Não faça**|
|:------------|:--------------|
|Use campos de texto para aceitar entrada de dados em um formulário ou em uma página.|Não use campos de texto para renderizar uma cópia básica como parte de um elemento de corpo de uma página.|
|Rotule os campos de texto com nomes úteis.|Não use campos de texto para entrada de data ou hora. Em vez disso, use um seletor de datetime.|
|Use texto de espaço reservado conciso para especificar qual conteúdo deve ser inserido.|Não use campos de texto se você puder predefinir as opções de entrada válidas. Em vez disso, use uma lista suspensa.|
|Forneça todos os estados apropriados para o campo de texto (static, hover, focus, engaged, unavailable, error).||
|Marque claramente os campos de texto obrigatórios e opcionais.||
|Sempre que possível, formate os campos de texto de acordo com o formato de dados esperado. Por exemplo, ao capturar um número de telefone de dez dígitos, use três campos separados para armazenar as diferentes partes do número de telefone.||

## <a name="variants"></a>Variantes

|**Variação**|**Descrição**|**Exemplo**|
|:------------|:--------------|:----------|
|**TextField padrão**|Use como o campo de texto padrão.|![Imagem de TextField padrão](../../images/textfieldDefault.png)<br/>|
|**TextField desativado**|Use quando o campo de texto estiver desativado.|![Imagem de TextField desativado](../../images/textfieldDisabled.png)<br/>|
|**TextField obrigatório**|Use quando a entrada do campo de texto for obrigatória.|![Imagem de TextField obrigatório](../../images/textfieldRequired.png)<br/>|
|**TextField com um espaço reservado**|Use quando um texto de espaço reservado for necessário.|![Imagem de TextField com um espaço reservado](../../images/textfieldPlaceholder.png)<br/>|
|**TextField com várias linhas**|Use quando muitas linhas de texto forem necessárias.|![Imagem de TextField com um espaço reservado](../../images/textfieldMulti.png)<br/>|

## <a name="implementation"></a>Implementação

Para saber mais, confira [TextField](https://dev.office.com/fabric#/components/textfield) e [Primeiros passos com exemplo de código do Fabric React](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact).

## <a name="additional-resources"></a>Recursos adicionais

- [Padrões de design da experiência do usuário](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)

- [Office UI Fabric em Suplementos do Office](office-ui-fabric.md)
