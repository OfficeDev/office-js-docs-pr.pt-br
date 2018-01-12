# <a name="dropdown-component-in-office-ui-fabric"></a>Componente DropDown no Office UI Fabric

Um menu suspenso é uma lista de opções que é mostrada clicando em um botão suspenso. Use uma lista ou menu suspensos para simplificar o design da interface do usuário e quando os usuários devem fazer uma escolha dentro da interface do usuário. Quando a lista colapsa, o item selecionado fica visível. Para alterar o item escolhido, os usuários abrem a lista e selecionam um novo valor.
  
#### <a name="example-drop-down-in-a-task-pane"></a>Exemplo: Menu suspenso em um painel de tarefas

<br/>

![Uma imagem mostrando o menu suspenso](../../images/overview_withApp_dropdown.png)

<br/>

## <a name="best-practices"></a>Práticas recomendadas

|**Faça**|**Não faça**|
|:------------|:--------------|
|Use um menu suspenso quando for mais provável que a opção selecionada seja selecionada do que outras opções. Por outro lado, ChoiceGroup ou botões de opção colocam a mesma ênfase em todas as opções exibidas.|Não use um menu suspenso quando todas as opções forem igualmente susceptíveis de serem selecionadas.|
|Use um menu suspenso quando houver múltiplas opções que possam ser colapsadas em um campo. Além disso, use um menu suspenso para listas longas de itens, ou quando o espaço na tela for restringido.|Não use um menu suspenso se houver menos de duas opções. Em vez disso, use uma caixa de seleção.|
|Use instruções ou palavras encurtadas em um menu suspenso.| |

## <a name="variants"></a>Variantes

|**Variação**|**Descrição**|**Exemplo**|
|:------------|:--------------|:----------|
|**Menu suspenso básico e não controlado **|Use quando houver muitas opções disponíveis para escolha.|![Imagem no menu suspenso básico e não controlado](../../images/dropdownUncontrolled.png)<br/>|
|**Menu suspenso não controlado e desabilitado com o defaultSelectedKey**|Estado desabilitado do menu suspenso.|![Menu suspenso não controlado e desabilitado com a imagem defaultSelectedKey](../../images/dropdownDisabled.png)<br/>|
|**Menu suspenso controlado**|Use quando o item selecionado padrão for influenciado por outra localização em sua interface de usuário, e o item selecionado no menu suspenso precisar ser mantido.|![Imagem do menu suspenso controlado](../../images/dropdownControlled.png)<br/>|

## <a name="implementation"></a>Implementação

Para saber mais, confira [Lista suspensa](https://dev.office.com/fabric#/components/dropdown) e [Primeiros passos com exemplo de código do Fabric React](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact).

## <a name="additional-resources"></a>Recursos adicionais

- [Padrões de design da experiência do usuário](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)

- [Office UI Fabric em Suplementos do Office](office-ui-fabric.md)
