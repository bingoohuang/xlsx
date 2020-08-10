package xlsx

import "strings"

// PlaceholderValue represents a placeholder value.
type PlaceholderValue struct {
	Content string

	Parts []PlaceholderPart
}

// HasPlaceholders tells that the PlaceholderValue has any placeholders.
func (p *PlaceholderValue) HasPlaceholders() bool {
	for _, p := range p.Parts {
		if p.Var != "" {
			return true
		}
	}

	return false
}

// Interpolate interpolates placeholders with vars.
func (p *PlaceholderValue) Interpolate(vars map[string]string) string {
	content := ""

	for _, p := range p.Parts {
		if p.Var != "" {
			content += vars[p.Var]
		} else {
			content += p.Part
		}
	}

	return content
}

// ParseVars parses the vars from the content.
func (p *PlaceholderValue) ParseVars(content string) (outVars map[string]string, matched bool) {
	outVars = make(map[string]string)

	for i := 0; i < len(p.Parts); i++ {
		v := p.Parts[i]

		if v.Var == "" {
			if !strings.HasPrefix(content, v.Part) {
				return nil, false
			}

			content = content[len(v.Part):]

			continue
		}

		if i+1 >= len(p.Parts) {
			outVars[v.Var] = content

			continue
		}

		i++
		v2 := p.Parts[i]
		v2Pos := strings.Index(content, v2.Part)

		if v2Pos < 0 {
			return nil, false
		}

		outVars[v.Var] = content[:v2Pos]
		content = content[v2Pos+len(v2.Part):]
	}

	return outVars, true
}

// PlaceholderPart is a placeholder sub Part after parsing.
type PlaceholderPart struct {
	Part string
	Var  string
}

// ParsePlaceholder parses placeholders in the content.
func ParsePlaceholder(content string) PlaceholderValue {
	pos := 0
	parts := make([]PlaceholderPart, 0)

	for {
		contentPos := content[pos:]
		lp := strings.Index(contentPos, "{{")

		if lp < 0 {
			if len(contentPos) > 0 {
				parts = append(parts, PlaceholderPart{
					Part: contentPos,
				})
			}

			break
		}

		rp := strings.Index(content[pos+lp:], "}}")
		if rp < 0 {
			if len(contentPos) > 0 {
				parts = append(parts, PlaceholderPart{
					Part: contentPos,
				})
			}

			break
		}

		if lp > 0 {
			parts = append(parts, PlaceholderPart{
				Part: contentPos[:lp],
			})
		}

		pl := content[pos+lp : pos+lp+rp+2]
		varName := strings.TrimSpace(pl[2 : len(pl)-2])

		parts = append(parts, PlaceholderPart{Part: pl, Var: varName})

		pos += lp + rp + 2 // nolint:gomnd
	}

	return PlaceholderValue{Content: content, Parts: parts}
}
