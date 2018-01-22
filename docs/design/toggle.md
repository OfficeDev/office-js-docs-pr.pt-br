# <a name="toggle-component-in-office-ui-fabric"></a>Componente de alternância no Office UI Fabric

As alternâncias representam uma opção física para ativar ou desativar recursos. Use alternâncias para apresentar duas opções mutuamente exclusivas (por exemplo, ativar ou desativar), em que a escolha de uma opção resulta em uma ação imediata.
  
#### <a name="example-toggle-in-a-task-pane"></a>Exemplo: Alternância em um painel de tarefas


![Uma imagem mostrando a alternância](../images/overview_withApp_toggle.png)

<br/>

## <a name="best-practices"></a>Práticas recomendadas

|**Faça**|**Não faça**|
|:------------|:--------------|
|Use alternâncias para configurações binárias quando as alterações são imediatamente aplicadas.<br/><br/>![Exemplo do que fazer com alternâncias](../images/toggleDo.png)<br/>|Não use alternâncias se os usuários tiverem que executar uma etapa adicional antes que as alterações entrem em vigor.<br/><br/>![Exemplo do que não fazer com alternâncias](../images/toggleDont.png)<br/>|
|Substitua os rótulos **Ativar** e **Desativar** somente se houver rótulos mais específicos a serem usados para uma configuração. Use rótulos curtos (três a quatro caracteres) que representem opostos binários.| |

## <a name="variants"></a>Variantes

|**Variação**|**Descrição**|**Exemplo**|
|:------------|:--------------|:----------|
|**Ativado e marcado**|Use quando o estado de alternância estiver ativo.|![Imagem de habilitado e marcado](../images/toggleEnabledOn.png)<br/>|
|**Habilitado e desmarcado**|Use quando o estado de alternância estiver inativo.|![Imagem de ativado e desmarcado](../images/toggleEnabledOff.png)<br/>|
|**Desabilitado e marcado**|Use quando o estado ativo não puder ser alterado.|![Imagem de desabilitado e marcado](../images/toggleDisabledOn.png)<br/>|
|**Desabilitado e desmarcado**|Use quando o estado inativo não puder ser alterado.|![Imagem de desabilitado e desmarcado](../images/toggleDisabledOff.png)<br/>|

## <a name="implementation"></a>Implementação

Para saber mais, confira [Alternância](https://dev.office.com/fabric#/components/toggle) e [Primeiros passos com exemplo de código do Fabric React](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact).

## <a name="additional-resources"></a>Recursos adicionais

- [Padrões de design da experiência do usuário](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)

- [Office UI Fabric em Suplementos do Office](office-ui-fabric.md)
