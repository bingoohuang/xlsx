package xlsx_test

import (
	"testing"

	"github.com/bingoohuang/xlsx"
	"github.com/stretchr/testify/assert"
)

func TestParsePlaceholder(t *testing.T) {
	assert.Equal(t, xlsx.PlaceholderValue{
		Content: "Age",
		Parts:   []xlsx.PlaceholderPart{{Part: "Age"}},
	}, xlsx.ParsePlaceholder("Age"))

	assert.Equal(t, xlsx.PlaceholderValue{
		Content: "{{name}}",
		Parts:   []xlsx.PlaceholderPart{{Part: "{{name}}", Var: "name"}},
	}, xlsx.ParsePlaceholder("{{name}}"))

	assert.Equal(t, xlsx.PlaceholderValue{
		Content: "{{name}} {{ age }}",
		Parts: []xlsx.PlaceholderPart{
			{Part: "{{name}}", Var: "name"},
			{Part: " ", Var: ""},
			{Part: "{{ age }}", Var: "age"},
		},
	}, xlsx.ParsePlaceholder("{{name}} {{ age }}"))

	assert.Equal(t, xlsx.PlaceholderValue{
		Content: "Hello {{name}} world {{ age }}",
		Parts: []xlsx.PlaceholderPart{
			{Part: "Hello ", Var: ""},
			{Part: "{{name}}", Var: "name"},
			{Part: " world ", Var: ""},
			{Part: "{{ age }}", Var: "age"},
		},
	}, xlsx.ParsePlaceholder("Hello {{name}} world {{ age }}"))

	assert.Equal(t, xlsx.PlaceholderValue{
		Content: "Age{{",
		Parts: []xlsx.PlaceholderPart{
			{Part: "Age{{", Var: ""},
		},
	}, xlsx.ParsePlaceholder("Age{{"))

	plName := xlsx.ParsePlaceholder("{{name}}")
	vars, ok := plName.ParseVars("bingoohuang")
	assert.True(t, ok)
	assert.Equal(t, map[string]string{"name": "bingoohuang"}, vars)

	nameArgs := xlsx.ParsePlaceholder("{{name}} {{ age }}")
	vars, ok = nameArgs.ParseVars("bingoohuang 100")
	assert.True(t, ok)
	assert.Equal(t, map[string]string{"name": "bingoohuang", "age": "100"}, vars)

	nameArgs = xlsx.ParsePlaceholder("中国{{v1}}人民{{v2}}")
	vars, ok = nameArgs.ParseVars("中国中央人民政府")
	assert.True(t, ok)
	assert.Equal(t, map[string]string{"v1": "中央", "v2": "政府"}, vars)
}
