package com.robin.rapidoffice.utils;

import java.io.IOException;

@FunctionalInterface
public interface ThrowableConsumer<T> {
    void accept(T t) throws IOException;
}
