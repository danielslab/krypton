/***************************************************************************
** Application: Krypton: ThinkSys Automation Generic System
** File Name: Extension.cs
** Copyright 2012, ThinkSys Software
** All rights reserved
** Developed by: ThinkSys Software 
** Description: Custom Extension Methods
*****************************************************************************/
using System;

namespace Krypton
{
    public static class Extension
    {
        public static bool IsNullOrWhiteSpace(this string stringVal)
        {
            return string.IsNullOrWhiteSpace(stringVal);
        }

        public static T[] RemoveAt<T>(this T[] source, int index)
        {
            T[] dest = new T[source.Length - 1];
            if (index > 0)
                Array.Copy(source, 0, dest, 0, index);
            if (index < source.Length - 1)
                Array.Copy(source, index + 1, dest, index, source.Length - index - 1);
            return dest;
        }

    }

}
