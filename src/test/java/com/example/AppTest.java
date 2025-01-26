package com.example;

import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;
import org.junit.Test;

/**
 * Unit test for simple App.
 */
public class AppTest {

    @Test
    public void testIsValidEmail() {
        assertTrue(Utils.isValidEmail("test@example.com"));
        assertFalse(Utils.isValidEmail("invalid-email"));
        assertFalse(Utils.isValidEmail(null));
    }

    @Test
    public void shouldAnswerWithTrue() {
        assertTrue(true);
    }
}
