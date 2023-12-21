package excelizemapper

type Format func(interface{}) string

type options struct {
	tagKey       string
	autoSort     bool
	defaultWidth float64
	formatterMap map[string]Format
}

type Option func(o *options)

// WithTagKey set tag key
//
// default is "excelize-mapper"
func WithTagKey(tagKey string) Option {
	return func(o *options) {
		o.tagKey = tagKey
	}
}

// WithFormatter set formatter
func WithFormatter(name string, format Format) Option {
	return func(o *options) {
		o.formatterMap[name] = format
	}
}

// WithAutoSort set auto sort
//
// if auto sort is false, use tag index. default is true.
func WithAutoSort(autoSort bool) Option {
	return func(o *options) {
		o.autoSort = autoSort
	}
}

// WithDefaultWidth set default width
func WithDefaultWidth(width float64) Option {
	return func(o *options) {
		o.defaultWidth = width
	}
}
