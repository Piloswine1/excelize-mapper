package excelizemapper

type options struct {
	TagKey       string
	FormatterMap map[string]func(interface{}) string
}

type Option func(o *options)

func WithTagKey(tagKey string) Option {
	return func(o *options) {
		o.TagKey = tagKey
	}
}

func WithFormatter(name string, format func(interface{}) string) Option {
	return func(o *options) {
		o.FormatterMap[name] = format
	}
}
