package it.firegloves.mempoi.domain.datatransformation;


import it.firegloves.mempoi.exception.MempoiException;

public class DataTransformationChain<I, O> {
    private final DataTransformationFunction<I, O> current;

    public DataTransformationChain(DataTransformationFunction<I, O> current) {
        this.current = current;
    }

    public <R> DataTransformationChain<I, R> chainUp(DataTransformationFunction<O, R> next) {
        return new DataTransformationChain<>(input -> next.transform(current.transform(input)));
    }

    public O execute(I input) throws MempoiException {
        return current.transform(input);
    }
}
