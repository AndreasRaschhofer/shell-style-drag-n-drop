using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace shell_style_drag_n_drop
{
    [ComVisible(true)]
    public class VirtualFileDataObject : IDataObject, IDisposable
    {
        // These are helper functions for managing STGMEDIUM structures

        [DllImport("urlmon.dll")]
        private static extern int CopyStgMedium(ref STGMEDIUM pcstgmedSrc, ref STGMEDIUM pstgmedDest);

        [DllImport("ole32.dll")]
        private static extern void ReleaseStgMedium(ref STGMEDIUM pmedium);

        private readonly List<(FORMATETC Format, STGMEDIUM Medium)> _storage;

        public VirtualFileDataObject() => _storage = new List<(FORMATETC, STGMEDIUM)>();

        ~VirtualFileDataObject() => Dispose(false);

        private void Dispose(bool disposing)
        {
            if (disposing)
            {
                // No managed objects to release
            }

            // Always release unmanaged objects
            ClearStorage();
        }

        public void Dispose()
        {
            Dispose(true);
        }

        private void ClearStorage()
        {
            foreach (var value in _storage)
            {
                STGMEDIUM medium = value.Medium;
                ReleaseStgMedium(ref medium);
            }

            _storage.Clear();
        }

        // COM Constants
        private const int OLE_E_ADVISENOTSUPPORTED = unchecked((int)0x80040003);

        private const int DV_E_FORMATETC = unchecked((int)0x80040064);
        private const int DV_E_TYMED = unchecked((int)0x80040069);
        private const int DV_E_CLIPFORMAT = unchecked((int)0x8004006A);
        private const int DV_E_DVASPECT = unchecked((int)0x8004006B);

        // Unsupported Functions
        int IDataObject.DAdvise(ref FORMATETC pFormatetc, ADVF advf, IAdviseSink adviseSink, out int connection)
            => throw Marshal.GetExceptionForHR(OLE_E_ADVISENOTSUPPORTED);

        void IDataObject.DUnadvise(int connection)
            => throw Marshal.GetExceptionForHR(OLE_E_ADVISENOTSUPPORTED);

        int IDataObject.EnumDAdvise(out IEnumSTATDATA enumAdvise)
            => throw Marshal.GetExceptionForHR(OLE_E_ADVISENOTSUPPORTED);

        int IDataObject.GetCanonicalFormatEtc(ref FORMATETC formatIn, out FORMATETC formatOut)
        {
            formatOut = formatIn;
            return DV_E_FORMATETC;
        }

        void IDataObject.GetDataHere(ref FORMATETC format, ref STGMEDIUM medium)
            => throw new NotSupportedException();

        IEnumFORMATETC IDataObject.EnumFormatEtc(DATADIR direction)
        {
            if (direction == DATADIR.DATADIR_GET)
                return new EnumFORMATETC(_storage);

            throw new NotImplementedException("OLE_S_USEREG");
        }

        /// <summary>
        /// Gets the specified data.
        /// </summary>
        /// <param name="format">The requested data format.</param>
        /// <param name="medium">When the function returns, contains the requested data.</param>
        public void GetData(ref FORMATETC format, out STGMEDIUM medium)
        {
            // Locate the data
            foreach (var pair in _storage)
            {
                if ((pair.Format.tymed & format.tymed) > 0
                    && pair.Format.dwAspect == format.dwAspect
                    && pair.Format.cfFormat == format.cfFormat)
                {
                    // Found it. Return a copy of the data.
                    STGMEDIUM source = pair.Medium;
                    medium = CopyMedium(ref source);
                    return;
                }
            }

            // Didn't find it. Return an empty data medium.
            medium = new STGMEDIUM();
        }

        /// <summary>
        /// Determines if data of the requested format is present.
        /// </summary>
        /// <param name="format">The request data format.</param>
        /// <returns>Returns the status of the request. If the data is present, S_OK is returned.
        /// If the data is not present, an error code with the best guess as to the reason is returned.</returns>
        public int QueryGetData(ref FORMATETC format)
        {
            // We only support CONTENT aspect
            if ((DVASPECT.DVASPECT_CONTENT & format.dwAspect) == 0)
                return DV_E_DVASPECT;

            int ret = DV_E_TYMED;

            // Try to locate the data
            // TODO: The ret, if not S_OK, is only relevant to the last item
            foreach (var pair in _storage)
            {
                if ((pair.Format.tymed & format.tymed) > 0)
                {
                    if (pair.Format.cfFormat == format.cfFormat)
                    {
                        // Found it, return S_OK;
                        return 0;
                    }
                    else
                    {
                        // Found the medium type, but wrong format
                        ret = DV_E_CLIPFORMAT;
                    }
                }
                else
                {
                    // Mismatch on medium type
                    ret = DV_E_TYMED;
                }
            }

            return ret;
        }

        /// <summary>
        /// Sets data in the specified format into storage.
        /// </summary>
        /// <param name="formatIn">The format of the data.</param>
        /// <param name="medium">The data.</param>
        /// <param name="release">If true, ownership of the medium's memory will be transferred
        /// to this object. If false, a copy of the medium will be created and maintained, and
        /// the caller is responsible for the memory of the medium it provided.</param>
        public void SetData(ref FORMATETC formatIn, ref STGMEDIUM medium, bool release)
        {
            // If the format exists in our storage, remove it prior to resetting it
            foreach (var pair in _storage)
            {
                if ((pair.Format.tymed & formatIn.tymed) > 0
                    && pair.Format.dwAspect == formatIn.dwAspect
                    && pair.Format.cfFormat == formatIn.cfFormat)
                {
                    _storage.Remove(pair);
                    break;
                }
            }

            // If release is true, we'll take ownership of the medium.
            // If not, we'll make a copy of it.
            STGMEDIUM sm = medium;
            if (!release)
                sm = CopyMedium(ref medium);

            // Add it to the internal storage
            _storage.Add((formatIn, sm));
        }

        /// <summary>
        /// Creates a copy of the STGMEDIUM structure.
        /// </summary>
        /// <param name="medium">The data to copy.</param>
        /// <returns>The copied data.</returns>
        private STGMEDIUM CopyMedium(ref STGMEDIUM medium)
        {
            STGMEDIUM sm = new STGMEDIUM();
            int hr = CopyStgMedium(ref medium, ref sm);
            if (hr != 0)
                throw Marshal.GetExceptionForHR(hr);

            return sm;

        }
    }
}
