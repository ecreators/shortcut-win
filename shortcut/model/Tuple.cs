namespace shortcut.model {
    public class Tuple<T1,T2> {


        public Tuple(T1 val1, T2 val2) {
            IsSet = val1;
            Value = val2;
        }

        public T1 IsSet { get; set; }

        public T2 Value { get; set; }
    }
}
