# <a name="defaultsettings-element"></a>Elemento DefaultSettings

Especifica o local de origem padrão e outras configurações padrão para seu suplemento de conteúdo ou de painel de tarefas.

**Tipo de suplemento:** Conteúdo, Painel de tarefas

## <a name="syntax"></a>Sintaxe

```XML
<DefaultSettings>
  ...
</DefaultSettings>
```

## <a name="contained-in"></a>Contido em

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>Pode conter

|**Element**|**Content**|**Mail**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[SourceLocation](sourcelocation.md)|x||x|
|[RequestedWidth](requestedwidth.md)|x|||
|[RequestedHeight](requestedheight.md)|x|||

## <a name="remarks"></a>Comentários

O local de origem e outras configurações no elemento **DefaultSettings** se aplicam apenas a suplementos de conteúdo e de painel de tarefas. Para suplementos de e-mail, você especifica os locais padrão para os arquivos de origem e outras configurações padrão no elemento [FormSettings](formsettings.md).

