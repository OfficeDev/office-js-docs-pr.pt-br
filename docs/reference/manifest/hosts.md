# <a name="hosts-element"></a>Elemento Hosts

Especifica o aplicativo cliente do Office no qual o suplemento do Office será ativado. Contém um conjunto de elementos **Host** e suas configurações. 

Quando incluído no nó [VersionOverrides](versionoverrides.md) , este elemento substitui o elemento **Hosts** na parte pai do manifesto. 

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  [Host](host.md)    |  Sim   |  Descreve um host e suas configurações. |
