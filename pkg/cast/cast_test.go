package cast_test

import (
	"reflect"
	"testing"
	"time"

	. "github.com/bingoohuang/xlsx/pkg/cast"

	"github.com/stretchr/testify/assert"
)

func TestPopulate(t *testing.T) {
	prop := map[string]string{
		"key1":        "value1",
		"key2":        "yes",
		"key3":        "true",
		"XingMing":    "kongrong",
		"foo-bar":     "foobar",
		"NI-HAO":      "10s",
		"ta-hao":      "10s",
		"ORDER_PRICE": "100",
		"order_items": "10",
		"HelloWorld":  "10",
	}

	type MySub2 struct {
		XingMing string
	}

	type MySub struct {
		Key1 string `prop:"key1"`
		Key2 bool
		Key3 *bool
	}

	type my struct {
		MySub
		*MySub2
		FooBar     *string
		NiHao      time.Duration
		TaHao      *time.Duration
		xx         string
		YY         string
		OrderPrice int
		OrderItems int
		HelloWorld *int
	}

	var (
		m my
		x int
	)

	it := assert.New(t)

	err := PopulateStruct(m, "prop", nil)
	it.Error(err)

	err = PopulateStruct(&x, "prop", nil)
	it.Error(err)

	mapGetter := func(m map[string]string) func(filedName, tagValue string) (interface{}, bool) {
		return func(filedName, tagValue string) (interface{}, bool) {
			return TryFind(filedName, tagValue, func(name string) (interface{}, bool) {
				v, ok := m[name]
				return v, ok
			})
		}
	}

	err = PopulateStruct(&m, "prop", mapGetter(prop))
	it.Nil(err)

	foobar := "foobar"
	HelloWorld := 10
	key3 := true
	taHao := 10 * time.Second

	it.Equal(my{
		MySub: MySub{
			Key1: "value1",
			Key2: true,
			Key3: &key3,
		},
		MySub2: &MySub2{
			XingMing: "kongrong",
		},
		FooBar:     &foobar,
		NiHao:      10 * time.Second,
		TaHao:      &taHao,
		xx:         "",
		YY:         "",
		OrderPrice: 100,
		OrderItems: 10,
		HelloWorld: &HelloWorld,
	}, m)

	prop = map[string]string{
		"NI-HAO":      "10x",
		"ta-hao":      "10s",
		"ORDER_PRICE": "100",
		"order_items": "10",
		"HelloWorld":  "10",
	}

	type Myx struct {
		NiHao time.Duration
	}

	type Myy struct {
		*Myx
	}

	var (
		myx Myx
		myy Myy
	)

	it.Error(PopulateStruct(&myx, "prop", mapGetter(prop)))
	it.Error(PopulateStruct(&myy, "prop", mapGetter(prop)))
}

func TestCastAny(t *testing.T) {
	type args struct {
		s string
		t reflect.Type
	}

	vTrue := true
	tests := []struct {
		name    string
		args    args
		want    reflect.Value
		wantErr bool
	}{
		{
			name:    "bool",
			args:    args{"yes", reflect.TypeOf(false)},
			want:    reflect.ValueOf(true),
			wantErr: false,
		},
		{
			name:    "*bool",
			args:    args{"yes", reflect.PtrTo(reflect.TypeOf(false))},
			want:    reflect.ValueOf(&vTrue),
			wantErr: false,
		},
		{
			name:    "bad bool",
			args:    args{"bad", reflect.TypeOf(false)},
			want:    InvalidValue,
			wantErr: true,
		},
	}

	for _, tt := range tests {
		tt := tt

		t.Run(tt.name, func(t *testing.T) {
			got, err := CastAny(tt.args.s, tt.args.t)
			if (err != nil) != tt.wantErr {
				t.Errorf("CastAny() error = %v, wantErr %v", err, tt.wantErr)
				return
			}
			if err != nil {
				if got != InvalidValue {
					t.Errorf("CastAny() got = %v, want %v", got, tt.want)
				}

				return
			}

			if !reflect.DeepEqual(got.Interface(), tt.want.Interface()) {
				t.Errorf("CastAny() got = %v, want %v", got, tt.want)
			}
		})
	}
}
