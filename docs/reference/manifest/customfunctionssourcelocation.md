# <a name="sourcelocation-element"></a>Elemento SourceLocation

Define a localização de um recurso necessário para os elementos Page ou Script usados por funções personalizadas no Excel.

## <a name="attributes"></a>Atributos

| **Atributo** | **Obrigatório** | **Descrição**                                                                      |
|---------------|--------------|--------------------------------------------------------------------------------------|
| resid         | Sim          | O nome de um recurso de URL definido na seção &lt;Recursos&gt; do manifesto. |

## <a name="child-elements"></a>Elementos filho

Nenhum

## <a name="example"></a>Exemplo

```xml
<SourceLocation resid="pageURL"/>
```