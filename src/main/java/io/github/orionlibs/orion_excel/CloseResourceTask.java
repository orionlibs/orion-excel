package io.github.orionlibs.orion_excel;

import java.io.Closeable;
import java.io.IOException;

class CloseResourceTask
{
    static boolean closeResource(Closeable closeable)
    {
        if(closeable != null)
        {
            try
            {
                closeable.close();
                return true;
            }
            catch(IOException e)
            {
                e.printStackTrace();
                return false;
            }
        }
        return false;
    }
}