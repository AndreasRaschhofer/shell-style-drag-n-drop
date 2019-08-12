using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading.Tasks;

namespace shell_style_drag_n_drop
{
    [ComVisible(true)]
    internal class EnumFORMATETC : IEnumFORMATETC
    {
        private readonly FORMATETC[] _formats;

        private int _currentIndex = 0;

        internal EnumFORMATETC(List<(FORMATETC Format, STGMEDIUM Medium)> storage)
        {
            // Get the formats from the list
            _formats = storage.Select(t => t.Format).ToArray();
        }

        private EnumFORMATETC(FORMATETC[] formats)
        {
            // Get the formats as a copy of the array
            _formats = new FORMATETC[formats.Length];
            formats.CopyTo(_formats, 0);
        }

        public void Clone(out IEnumFORMATETC newEnum)
        {
            var clone = new EnumFORMATETC(_formats)
            {
                _currentIndex = _currentIndex
            };

            newEnum = clone;
        }

        private const int S_OK = 0;
        private const int S_FALSE = 1;

        /// <summary>
        /// Retrieves the next elements from the enumeration.
        /// </summary>
        /// <param name="celt">The number of elements to retrieve.</param>
        /// <param name="rgelt">An array to receive the formats requested.</param>
        /// <param name="pceltFetched">An array to receive the number of element fetched.</param>
        /// <returns>If the fetched number of formats is the same as the requested number, S_OK is returned.
        /// There are several reasons S_FALSE may be returned:
        ///     (1) The requested number of elements is less than or equal to zero.
        ///     (2) The rgelt parameter equals null.
        ///     (3) There are no more elements to enumerate.
        ///     (4) The requested number of elements is greater than one and pceltFetched equals null or does not
        /// have at least one element in it. 
        ///     (5) The number of fetched elements is less than the number of requested elements.
        /// </returns>
        public int Next(int celt, FORMATETC[] rgelt, int[] pceltFetched)
        {
            // Start with zero fetched, in case we return early
            if (pceltFetched != null && pceltFetched.Length > 0)
                pceltFetched[0] = 0;

            // This will count down as we fetch elements
            int cReturn = celt;

            // Short circuit if they didn't request any elements, or didn't
            // provide room in the return array, or there are not more elements to enumerate.
            if (celt <= 0 || rgelt == null || _currentIndex >= _formats.Length)
                return S_FALSE;

            // If the number of requested elements is not one, then we must
            // be able to tell the caller how many elements were fetched.
            if ((pceltFetched == null || pceltFetched.Length < 1) && celt != 1)
                return S_FALSE;

            // If the number of elements in the return array is too small, we throw.
            // This is not a likely scenario, hence the exception.
            if (rgelt.Length < celt)
                throw new ArgumentException("The number of elements in the return array is less than the number of elements requested");

            // Fetch the elements.
            for (int i = 0; _currentIndex < _formats.Length && cReturn > 0; i++, cReturn--, _currentIndex++)
                rgelt[i] = _formats[_currentIndex];

            // Return the number of elements fetched
            if (pceltFetched != null && pceltFetched.Length > 0)
                pceltFetched[0] = celt - cReturn;

            // cReturn has the number of elements requested but not fetched.
            // It will be greater than zero, if multiple elements were requested
            // but we hit the end of the enumeration.
            return (cReturn == 0) ? S_OK : S_FALSE;
        }

        public int Reset()
        {
            _currentIndex = 0;
            return S_OK;
        }

        /// <summary>
        /// Skips the number of elements requested.
        /// </summary>
        /// <param name="celt">The number of elements to skip.</param>
        /// <returns>If there are not enough remaining elements to skip, returns S_FALSE. Otherwise, S_OK is returned.</returns>
        public int Skip(int celt)
        {
            if (_currentIndex + celt > _formats.Length)

                return S_FALSE;

            _currentIndex += celt;

            return S_OK;
        }
    }
}
