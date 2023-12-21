package excelizemapper

type options struct {
	tagKey       string
	autoSort     bool
	formatterMap map[string]func(interface{}) string
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
func WithFormatter(name string, format func(interface{}) string) Option {
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
