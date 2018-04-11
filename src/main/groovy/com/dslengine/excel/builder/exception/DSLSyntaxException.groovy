package com.dslengine.excel.builder.exception

class DSLSyntaxException extends RuntimeException {

    public DSLSyntaxException() {
        super()
    }

    public DSLSyntaxException(Throwable cause) {
        super(cause)
    }

    public DSLSyntaxException(String message) {
        super(message)
    }


    public DSLSyntaxException(String message, Throwable cause) {
        super(message, cause)
    }

}
