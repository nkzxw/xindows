
#include "stdafx.h"
#include "Image.h"

#include "Dither.h"

extern "C"
{
#include "jpglib6a/jinclude.h"
#define JPEG_INTERNALS
#include "jpglib6a/jpeglib.h"
#include "jpglib6a/jerror.h"
}

#define EXCEPTION_JPGLIB    0x1

class CImgTaskJpg : public CImgTask
{
public:
    DECLARE_MEMCLEAR_NEW_DELETE()

    virtual void Decode(BOOL* pfNonProgressive);
    virtual void BltDib(HDC hdc, RECT* prcDst, RECT* prcSrc, DWORD dwRop, DWORD dwFlags);

private:
    CImgBitsDIB* CreateImgBits();
};

/*
* Error exit handler: must not return to caller.
*
* Applications may override this if they want to get control back after
* an error.  Typically one would longjmp somewhere instead of exiting.
* The setjmp buffer can be made a private field within an expanded error
* handler object.  Note that the info needed to generate an error message
* is stored in the error object, so you can generate the message now or
* later, at your convenience.
* You should make sure that the JPEG object is cleaned up (with jpeg_abort
* or jpeg_destroy) at some point.
*/
EXTERN_C METHODDEF(void) error_exit(j_common_ptr cinfo)
{
    /* Let the memory manager delete any temp files before we die */
    jpeg_destroy(cinfo);
}

/*
* Actual output of an error or trace message.
* Applications may override this method to send JPEG messages somewhere
* other than stderr.
*/
EXTERN_C METHODDEF(void) output_message(j_common_ptr cinfo)
{
}

/*
* Decide whether to emit a trace or warning message.
* msg_level is one of:
*   -1: recoverable corrupt-data warning, may want to abort.
*    0: important advisory messages (always display to user).
*    1: first level of tracing detail.
*    2,3,...: successively more detailed tracing messages.
* An application might override this method if it wanted to abort on warnings
* or change the policy about which messages to display.
*/
EXTERN_C METHODDEF(void) emit_message(j_common_ptr cinfo, int msg_level)
{
}

/*
* Format a message string for the most recent JPEG error or message.
* The message is stored into buffer, which should be at least JMSG_LENGTH_MAX
* characters.  Note that no '\n' character is added to the string.
* Few applications should need to override this method.
*/
EXTERN_C METHODDEF(void) format_message (j_common_ptr cinfo, char* buffer)
{
}

/*
* Reset error state variables at start of a new image.
* This is called during compression startup to reset trace/error
* processing to default state, without losing any application-specific
* method pointers.  An application might possibly want to override
* this method if it has additional error processing state.
*/
EXTERN_C METHODDEF(void) reset_error_mgr(j_common_ptr cinfo)
{
    cinfo->err->num_warnings = 0;
    /* trace_level is not reset since it is an application-supplied parameter */
    cinfo->err->msg_code = 0; /* may be useful as a flag for "no error" */
}

/*
* Fill in the standard error-handling methods in a jpeg_error_mgr object.
* Typical call is:
*  struct jpeg_compress_struct cinfo;
*  struct jpeg_error_mgr err;
*
*  cinfo.err = jpeg_std_error(&err);
* after which the application may override some of the methods.
*/
EXTERN_C GLOBAL(struct jpeg_error_mgr*) jpeg_std_error (struct jpeg_error_mgr* err)
{
    err->error_exit = error_exit;
    err->emit_message = emit_message;
    err->output_message = output_message;
    err->format_message = format_message;
    err->reset_error_mgr = reset_error_mgr;

    err->trace_level = 0;     /* default = no tracing */
    err->num_warnings = 0;    /* no warnings emitted yet */
    err->msg_code = 0;        /* may be useful as a flag for "no error" */

    /* Initialize message table pointers */
    err->jpeg_message_table = NULL;
    err->last_jpeg_message = (int)JMSG_LASTMSGCODE - 1;

    err->addon_message_table = NULL;
    err->first_addon_message = 0; /* for safety */
    err->last_addon_message = 0;

    return err;
}

/* Expanded data source object for stdio input */
typedef struct
{
    struct jpeg_source_mgr pub; /* public fields */

    CImgTask* pImgTask;
    FILE* infile;               /* source stream */
    JOCTET* buffer;             /* start of buffer */
    boolean start_of_file;      /* have we gotten any data yet? */
} my_source_mgr;

typedef my_source_mgr*      my_src_ptr;

#define INPUT_BUF_SIZE      4096    /* choose an efficiently fread'able size */
#define ImgTask_BUF_SIZE    512     /* choose a size to allow overlap */

/*
* Initialize source --- called by jpeg_read_header
* before any data is actually read.
*/
EXTERN_C METHODDEF(void) init_source(j_decompress_ptr cinfo)
{
    my_src_ptr src = (my_src_ptr)cinfo->src;

    /* We reset the empty-input-file flag for each image,
    * but we don't clear the input buffer.
    * This is correct behavior for reading a series of images from one source.
    */
    src->start_of_file = TRUE;
}

/*
* Fill the input buffer --- called whenever buffer is emptied.
*
* In typical applications, this should read fresh data into the buffer
* (ignoring the current state of next_input_byte & bytes_in_buffer),
* reset the pointer & count to the start of the buffer, and return TRUE
* indicating that the buffer has been reloaded.  It is not necessary to
* fill the buffer entirely, only to obtain at least one more byte.
*
* There is no such thing as an EOF return.  If the end of the file has been
* reached, the routine has a choice of ERREXIT() or inserting fake data into
* the buffer.  In most cases, generating a warning message and inserting a
* fake EOI marker is the best course of action --- this will allow the
* decompressor to output however much of the image is there.  However,
* the resulting error message is misleading if the real problem is an empty
* input file, so we handle that case specially.
*
* In applications that need to be able to suspend compression due to input
* not being available yet, a FALSE return indicates that no more data can be
* obtained right now, but more may be forthcoming later.  In this situation,
* the decompressor will return to its caller (with an indication of the
* number of scanlines it has read, if any).  The application should resume
* decompression after it has loaded more data into the input buffer.  Note
* that there are substantial restrictions on the use of suspension --- see
* the documentation.
*
* When suspending, the decompressor will back up to a convenient restart point
* (typically the start of the current MCU). next_input_byte & bytes_in_buffer
* indicate where the restart point will be if the current call returns FALSE.
* Data beyond this point must be rescanned after resumption, so move it to
* the front of the buffer rather than discarding it.
*/
EXTERN_C METHODDEF(boolean) fill_input_buffer(j_decompress_ptr cinfo)
{
    my_src_ptr src = (my_src_ptr) cinfo->src;
    ULONG nbytes;

    if(!src->pImgTask->Read(src->buffer, ImgTask_BUF_SIZE, &nbytes))
    {
        if(src->start_of_file) /* Treat empty input file as fatal error */
        {
            ERREXIT(cinfo, JERR_INPUT_EMPTY);
        }

        WARNMS(cinfo, JWRN_JPEG_EOF);
        /* Insert a fake EOI marker */
        src->buffer[0] = (JOCTET)0xFF;
        src->buffer[1] = (JOCTET)JPEG_EOI;
        nbytes = 2;

        src->pub.next_input_byte = src->buffer;
        src->pub.bytes_in_buffer = nbytes;
        src->start_of_file = FALSE;

        return FALSE;
    }

    src->pub.next_input_byte = src->buffer;
    src->pub.bytes_in_buffer = nbytes;
    src->start_of_file = FALSE;

    return TRUE;
}

/*
* Skip data --- used to skip over a potentially large amount of
* uninteresting data (such as an APPn marker).
*
* Writers of suspendable-input applications must note that skip_input_data
* is not granted the right to give a suspension return.  If the skip extends
* beyond the data currently in the buffer, the buffer can be marked empty so
* that the next read will cause a fill_input_buffer call that can suspend.
* Arranging for additional bytes to be discarded before reloading the input
* buffer is the application writer's problem.
*/
EXTERN_C METHODDEF(void) skip_input_data(j_decompress_ptr cinfo, long num_bytes)
{
    my_src_ptr src = (my_src_ptr)cinfo->src;

    /* Just a dumb implementation for now.  Could use fseek() except
    * it doesn't work on pipes.  Not clear that being smart is worth
    * any trouble anyway --- large skips are infrequent.
    */
    if(num_bytes > 0)
    {
        while(num_bytes > (long)src->pub.bytes_in_buffer)
        {
            num_bytes -= (long)src->pub.bytes_in_buffer;
            (void)fill_input_buffer(cinfo);
        }
        src->pub.next_input_byte += (size_t)num_bytes;
        src->pub.bytes_in_buffer -= (size_t)num_bytes;
    }
}

/*
* An additional method that can be provided by data source modules is the
* resync_to_restart method for error recovery in the presence of RST markers.
* For the moment, this source module just uses the default resync method
* provided by the JPEG library.  That method assumes that no backtracking
* is possible.
*/


/*
* Terminate source --- called by jpeg_finish_decompress
* after all data has been read.  Often a no-op.
*
* NB: *not* called by jpeg_abort or jpeg_destroy; surrounding
* application must deal with any cleanup that should happen even
* for error exit.
*/
EXTERN_C METHODDEF(void) term_source(j_decompress_ptr cinfo)
{
    /* no work necessary here */
}

EXTERN_C GLOBAL(void) jpeg_ImgTask_src(j_decompress_ptr cinfo, CImgTask* pImgTask)
{
    my_src_ptr src;

    /* The source object and input buffer are made permanent so that a series
    * of JPEG images can be read from the same file by calling jpeg_stdio_src
    * only before the first one.  (If we discarded the buffer at the end of
    * one image, we'd likely lose the start of the next one.)
    * This makes it unsafe to use this manager and a different source
    * manager serially with the same JPEG object.  Caveat programmer.
    */
    if(cinfo->src == NULL)
    {
        /* first time for this JPEG object? */
        cinfo->src = (struct jpeg_source_mgr*)
            (*cinfo->mem->alloc_small)((j_common_ptr)cinfo, JPOOL_PERMANENT,
            SIZEOF(my_source_mgr));
        src = (my_src_ptr)cinfo->src;
        src->buffer = (JOCTET*)
            (*cinfo->mem->alloc_small)((j_common_ptr)cinfo, JPOOL_PERMANENT,
            ImgTask_BUF_SIZE * J_SIZEOF(JOCTET));
    }

    src = (my_src_ptr)cinfo->src;
    src->pub.init_source = init_source;
    src->pub.fill_input_buffer = fill_input_buffer;
    src->pub.skip_input_data = skip_input_data;
    src->pub.resync_to_restart = jpeg_resync_to_restart; /* use default method */
    src->pub.term_source = term_source;
    src->infile = NULL;
    src->pImgTask = pImgTask;
    src->pub.bytes_in_buffer = 0; /* forces fill_input_buffer on first read */
    src->pub.next_input_byte = NULL; /* until buffer loaded */
}

/*
* ERROR HANDLING:
*
* The JPEG library's standard error handler (jerror.c) is divided into
* several "methods" which you can override individually.  This lets you
* adjust the behavior without duplicating a lot of code, which you might
* have to update with each future release.
*
* Our example here shows how to override the "error_exit" method so that
* control is returned to the library's caller when a fatal error occurs,
* rather than calling exit() as the standard error_exit method does.
*
* We use C's setjmp/longjmp facility to return control.  This means that the
* routine which calls the JPEG library must first execute a setjmp() call to
* establish the return point.  We want the replacement error_exit to do a
* longjmp().  But we need to make the setjmp buffer accessible to the
* error_exit routine.  To do this, we make a private extension of the
* standard JPEG error handler object.  (If we were using C++, we'd say we
* were making a subclass of the regular error handler.)
*
* Here's the extended error handler struct:
*/
struct my_error_mgr
{
    struct jpeg_error_mgr pub; /* "public" fields */
};

typedef struct my_error_mgr*    my_error_ptr;

int x_MapGraysToVGAPalette[3] =
{
    0,
    7,
    15
};

/*
* Here's the routine that will replace the standard error_exit method:
*/
EXTERN_C METHODDEF(void) my_error_exit(j_common_ptr cinfo)
{
    /* Always display the message. */
    /* We could postpone this until after returning, if we chose. */
    (*cinfo->err->output_message)(cinfo);

    /* Return control to the setjmp point */
    RaiseException(EXCEPTION_JPGLIB, EXCEPTION_NONCONTINUABLE, 0, NULL);
}

/*
* Sample routine for JPEG decompression.  We assume that the JPEG file image
* is passed in.  We want to return a pointer on success, NULL on error.
*/
/* This version of the routine uses the IJG dithering code to dither into our 6x6x6 cube */
void CImgTaskJpg::Decode(BOOL* pfNonProgressive)
{
    /* This struct contains the JPEG decompression parameters and pointers to
    * working space (which is allocated as needed by the JPEG library).
    */
    struct jpeg_decompress_struct cinfo;
    /* We use our private extension JPEG error handler. */
    struct my_error_mgr jerr;
    /* More stuff */
    JSAMPARRAY buffer;  /* Output row buffer */
    int row_stride;     /* physical row width in output buffer */

    unsigned char HUGEP* pCurRow;

    int xsize, ysize;
    int irow;
    int x;
    int y;
    int num_rows_read;
    int notifyRow = 0;
    BYTE HUGEP* pbBits;
    int cbRow;
    BYTE xPixel;

    ERRBUF *pErrBuf1 = NULL, *pErrBuf2 = NULL;

    /* Step 1: allocate and initialize JPEG decompression object */

    /* We set up the normal JPEG error routines, then override error_exit. */
    cinfo.err = jpeg_std_error(&jerr.pub);
    jerr.pub.error_exit = my_error_exit;

    /* Establish the setjmp return context for my_error_exit to use. */
    __try
    {
        /* Now we can initialize the JPEG decompression object. */
        jpeg_create_decompress(&cinfo);

        /* Step 2: specify data source (eg, a file, or a memory buffer) */
        jpeg_ImgTask_src(&cinfo, this);

        /* Step 3: read file parameters with jpeg_read_header() */
        (void)jpeg_read_header(&cinfo, TRUE);

        /* We can ignore the return value from jpeg_read_header since
        *   (a) suspension is not possible with the stdio data source, and
        *   (b) we passed TRUE to reject a tables-only JPEG file as an error.
        * See libjpeg.doc for more info.
        */

        /* Step 4: set parameters for decompression */
        cinfo.dct_method = JDCT_ISLOW;

        switch(cinfo.jpeg_color_space)
        {
        case JCS_GRAYSCALE:
            if(_colorMode == 4)
            {
                cinfo.out_color_space = JCS_GRAYSCALE;
                cinfo.quantize_colors = TRUE;
                cinfo.desired_number_of_colors = 3;
                cinfo.two_pass_quantize = FALSE;
                cinfo.dither_mode = JDITHER_FS;
            }
            else
            {
                cinfo.out_color_space = JCS_GRAYSCALE;
                /* We want the actual RGB data here */
                cinfo.quantize_colors = FALSE;
            }
            break;

        default:
            if(_colorMode == 4)
            {
                cinfo.out_color_space = JCS_RGB;
                cinfo.quantize_colors = TRUE;
                cinfo.desired_number_of_colors = 16;
                cinfo.two_pass_quantize = FALSE;
                cinfo.dither_mode = JDITHER_FS;
                cinfo.colormap = (*cinfo.mem->alloc_sarray)
                    ((j_common_ptr)&cinfo, JPOOL_IMAGE, 16, 3);
                {
                    int i;

                    for(i=0; i<16; i++)
                    {
                        cinfo.colormap[RGB_RED][i]   = g_peVga[i].peRed;
                        cinfo.colormap[RGB_GREEN][i] = g_peVga[i].peGreen;
                        cinfo.colormap[RGB_BLUE][i]  = g_peVga[i].peBlue;
                    }
                }
                cinfo.actual_number_of_colors = 16;
            }
            else
            {
                cinfo.out_color_space = JCS_RGB;
                /* We want the actual RGB data here */
                cinfo.quantize_colors = FALSE;
            }
            break;
        }

        /* Step 5: Start decompressor */
        jpeg_start_decompress(&cinfo);

        /* We may need to do some setup of our own at this point before reading
        * the data.  After jpeg_start_decompress() we have the correct scaled
        * output image dimensions available, as well as the output colormap
        * if we asked for color quantization.
        * In this example, we need to make an output work buffer of the right size.
        */
        _xWid = xsize = cinfo.output_width;
        _yHei = ysize = cinfo.output_height;
        OnSize(_xWid, _yHei, -1);

        {
            CImgBitsDIB* pibd;
            _pImgBits = pibd = CreateImgBits();
            if(!pibd)
            {
                my_error_exit((j_common_ptr)&cinfo);
            }

            HRESULT hr = pibd->AllocDIB((_colorMode>24)?24:_colorMode, _xWid, _yHei, NULL, 0, -1, TRUE);
            if(hr)
            {
                my_error_exit((j_common_ptr)&cinfo);
            }

            pbBits = (BYTE*)pibd->GetBits();
            cbRow = pibd->CbLine();
            pibd->SetValidLines(0);
        }

        // Prepare for dithering, if necessary
        if(_colorMode<24 && _colorMode!=4)
        {
            if(FAILED(AllocDitherBuffers(_xWid, &pErrBuf1, &pErrBuf2)))
            {
                my_error_exit((j_common_ptr)&cinfo);
            }
        }

        /* JSAMPLEs per row in output buffer */
        row_stride = cinfo.output_width * cinfo.output_components;
        /* Make a sample array that will go away when done with image */
        buffer = (*cinfo.mem->alloc_sarray)
            ((j_common_ptr)&cinfo, JPOOL_IMAGE, row_stride, 8);

        /* Step 6: while (scan lines remain to be read) */
        /*           jpeg_read_scanlines(...); */
        y = 0;
        while(y < ysize)
        {
            num_rows_read = jpeg_read_scanlines(&cinfo, buffer, 8);

            if(!num_rows_read)
            {
                goto abort;
            }
            if(_colorMode == 4)
            {
                if(cinfo.out_color_space == JCS_RGB)
                {
                    for(irow=0; irow<num_rows_read; irow++)
                    {
                        pCurRow = pbBits + (long)cbRow*(long)(ysize-y-1); /* the DIB is stored upside down */

                        for(x=0; x<xsize; x+=2)
                        {
                            *pCurRow++ = (buffer[irow][x]<<4) | (buffer[irow][x+1]);
                        }
                        y++;
                    }
                }
                else
                {
                    AssertSz((cinfo.out_color_space==JCS_GRAYSCALE), "Illegal color space");
                    for(irow=0; irow<num_rows_read; irow++)
                    {
                        pCurRow = pbBits + (long)cbRow*(long)(ysize-y-1); /* the DIB is stored upside down */

                        for(x=0; x<xsize; x+=2)
                        {
                            *pCurRow++ = (x_MapGraysToVGAPalette[buffer[irow][x]]<<4)
                                | x_MapGraysToVGAPalette[buffer[irow][x+1]];
                        }
                        y++;
                    }
                }
            }
            else if(_colorMode == 8)
            {
                pCurRow = pbBits + (long)cbRow*(long)(ysize-y-1); /* the DIB is stored upside down */

                if(cinfo.out_color_space == JCS_RGB)
                {
                    Dith24rto8(pCurRow, &buffer[0][0], -cbRow, row_stride,  g_rgbHalftone, g_pInvCMAP, 
                        pErrBuf1, pErrBuf2, 0, xsize, y, num_rows_read);
                }
                else
                {
                    DithGray8to8(pCurRow, &buffer[0][0], -cbRow, row_stride, g_rgbHalftone, g_pInvCMAP, 
                        pErrBuf1, pErrBuf2, 0, xsize, y, num_rows_read);
                }

                y += num_rows_read;
            }
            else if(_colorMode==15 || _colorMode==16)
            {
                pCurRow = pbBits + (long)cbRow*(long) (ysize-y-1); /* the DIB is stored upside down */

                if(cinfo.out_color_space == JCS_RGB)
                {
                    if(_colorMode == 15)                
                    {
                        Convert24rto15((WORD*)pCurRow, &buffer[0][0], -cbRow, row_stride, 
                            0, xsize, y, num_rows_read);
                    }
                    else
                    {
                        Convert24rto16((WORD*)pCurRow, &buffer[0][0], -cbRow, row_stride,  
                            0, xsize, y, num_rows_read);
                    }
                }
                else
                {
                    if(_colorMode == 15)
                    {
                        DithGray8to15((WORD*)pCurRow, &buffer[0][0], -cbRow, row_stride, pErrBuf1, pErrBuf2, 
                            0, xsize, y, num_rows_read);
                    }
                    else
                    {
                        DithGray8to16((WORD*)pCurRow, &buffer[0][0], -cbRow, row_stride, pErrBuf1, pErrBuf2, 
                            0, xsize, y, num_rows_read);
                    }
                }

                y += num_rows_read;
            }
            else
            {
                if(cinfo.out_color_space == JCS_RGB)
                {
                    for(irow=0; irow<num_rows_read; irow++)
                    {
                        pCurRow = pbBits + (long)cbRow*(long)(ysize-y-1); /* the DIB is stored upside down */

                        for(x=0; x<xsize; x++)
                        {
                            /*
                            DIB's are stored blue-green-red (backwards)
                            */
                            *pCurRow++ = buffer[irow][x*3+RGB_BLUE];
                            *pCurRow++ = buffer[irow][x*3+RGB_GREEN];
                            *pCurRow++ = buffer[irow][x*3+RGB_RED];
                        }
                        y++;
                    }
                }
                else
                {
                    AssertSz((cinfo.out_color_space==JCS_GRAYSCALE), "Illegal color space");
                    for(irow=0; irow<num_rows_read; irow++)
                    {
                        pCurRow = pbBits + (long)cbRow*(long)(ysize-y-1); /* the DIB is stored upside down */

                        for(x=0; x<xsize; x++)
                        {
                            xPixel = buffer[irow][x];
                            *pCurRow++ = xPixel;
                            *pCurRow++ = xPixel;
                            *pCurRow++ = xPixel;
                        }
                        y++;
                    }
                }
            }

            _yBot = y - 1;

            ((CImgBitsDIB*)_pImgBits)->SetValidLines(_yBot+1);

            if(_yBot-notifyRow >= 4)
            {
                notifyRow = _yBot;
                OnProg(FALSE, IMGBITS_PARTIAL, FALSE, _yBot);
            }
        }

        _ySrcBot = -1;
        OnProg(TRUE, IMGBITS_TOTAL, FALSE, _yBot);

        /* Step 7: Finish decompression */
        (void)jpeg_finish_decompress(&cinfo);
        /* We can ignore the return value since suspension is not possible
        * with the stdio data source.
        */

        /* Step 8: Release JPEG decompression object */

        /* This is an important step since it will release a good deal of memory. */
        jpeg_destroy_decompress(&cinfo);

        /* After finish_decompress, we can close the input file.
        * Here we postpone it until after no more JPEG errors are possible,
        * so as to simplify the setjmp error logic above.  (Actually, I don't
        * think that jpeg_destroy can do an error exit, but why assume anything...)
        */

        /* At this point you may want to check to see whether any corrupt-data
        * warnings occurred (test whether jerr.pub.num_warnings is nonzero).
        */
        ((CImgBitsDIB*)_pImgBits)->SetValidLines(_yHei);

        if(pErrBuf1)
        {
            FreeDitherBuffers(pErrBuf1, pErrBuf2);
        }

        /* And we're done! */
        return;

abort:
        jpeg_destroy_decompress(&cinfo);
        if(pErrBuf1)
        {
            FreeDitherBuffers(pErrBuf1, pErrBuf2);
        }

        if(_yBot > 31)
        {
            _ySrcBot = _yBot + 1;
        }
        return;
    }
    __except(GetExceptionCode()==EXCEPTION_JPGLIB ? EXCEPTION_EXECUTE_HANDLER : EXCEPTION_CONTINUE_SEARCH)
    {
        /* If we get here, the JPEG code has signaled an error.
        * We need to clean up the JPEG object, close the input file, and return.
        */

        /*
        TODO call WAIT_Pop ?
        */
        jpeg_destroy_decompress(&cinfo);

        if(pErrBuf1)
        {
            FreeDitherBuffers(pErrBuf1, pErrBuf2);
        }

        return;
    }
}

/*
* SOME FINE POINTS:
*
* We cheated a bit by calling alloc_sarray() after jpeg_start_decompress();
* we should have done it beforehand to ensure that the space would be
* counted against the JPEG max_memory setting.  In some systems the above
* code would risk an out-of-memory error.  However, in general we don't
* know the output image dimensions before jpeg_start_decompress(), unless we
* call jpeg_calc_output_dimensions().  See libjpeg.doc for more about this.
*
* Scanlines are returned in the same order as they appear in the JPEG file,
* which is standardly top-to-bottom.  If you must emit data bottom-to-top,
* you can use one of the virtual arrays provided by the JPEG memory manager
* to invert the data.  See wrbmp.c for an example.
*
* As with compression, some operating modes may require temporary files.
* On some systems you may need to set up a signal handler to ensure that
* temporary files are deleted if the program is interrupted.  See libjpeg.doc.
*/

//  Performs a StretchDIBits for progressive draw (deals with
//  only some of the data being available etc
void CImgTaskJpg::BltDib(HDC hdc, RECT* prcDst, RECT* prcSrc, DWORD dwRop, DWORD dwFlags)
{
    _pImgBits->StretchBlt(hdc, prcDst, prcSrc, dwRop, dwFlags);
}

CImgBitsDIB* CImgTaskJpg::CreateImgBits()
{
    return new CImgBitsDIB;
}

CImgTask* NewImgTaskJpg()
{
    return new CImgTaskJpg;
}

// Detecting MMX --------------------------------------------------------------
extern "C" int IsMMX() // does the processor I'm running have MMX(tm) technology?
{
    volatile int iResult;
    HKEY    hkey = NULL;
    BOOL    fRet = TRUE;
    BOOL    fDefaultOverridden = FALSE;

    if(RegOpenKeyEx(HKEY_LOCAL_MACHINE, TEXT("Software\\Microsoft\\Internet Explorer"), 0, KEY_QUERY_VALUE, &hkey) == ERROR_SUCCESS)
    {
        BYTE    szData[5];
        DWORD   dwSize = sizeof(szData);
        DWORD   dwType;

        if(RegQueryValueEx(hkey, TEXT("UseMMX"), NULL, &dwType, szData, &dwSize) == ERROR_SUCCESS)
        {
            if(*szData)
            {
                fRet = (*szData != '0');
                fDefaultOverridden = TRUE;
            }
        }

        CloseHandle(hkey);
    }

    if(fDefaultOverridden && !fRet)
    {
        return fRet;
    }

    {
        typedef BOOL (WINAPI *PFNPFP)(DWORD dw);
        PFNPFP pfnIsPFP = (PFNPFP)GetProcAddress(GetModuleHandleA("kernel32"), "IsProcessorFeaturePresent");

        // On NT, we can just ask the OS if MMX instructions are available
        if(pfnIsPFP && !pfnIsPFP(PF_MMX_INSTRUCTIONS_AVAILABLE))
        {
            return(FALSE);
        }
    }

    // Just to be sure, try executing an MMX instruction and see if it barfs
    iResult = 1;

    __try
    {
        _asm    punpckldq   mm7,mm2
        _asm    emms
    }
    __except(EXCEPTION_EXECUTE_HANDLER)
    {
        iResult = 0;
    }

    // If the processor is a Pentium 3, disable MMX usage, unless the default action has been
    // overridden by the registry
    // Checkout: http://developer.intel.com/design/pentiumii/applnots/241618.htm
    if(!fDefaultOverridden && iResult)
    {
        // Bit masks for the CPUID information
#define CPU_TYPE        0x00003000
#define CPU_FAMILY      0x00000F00
#define CPU_MODEL       0x000000F0
#define CPU_STEPPING    0x0000000F

#define CPU_PENTIUM3    0x00000670      // Family = 6, Model = 7, Stepping doesn't matter
#define CPUID           {   __asm _emit 0x0f    __asm _emit 0xa2  }

        DWORD   dwCPUInfo;

        __asm   mov     eax, 1
            CPUID
            __asm   mov     dwCPUInfo, eax

            if((dwCPUInfo&CPU_PENTIUM3) == CPU_PENTIUM3)
            {
                iResult = 0;
            }

    }

    return iResult;
}