include config.mk

SRC = doctxt.c util.c miniz.c
OBJ = ${SRC:.c=.o}

MD2DOCX_SRC = md2docx.c util.c miniz.c md4c.c
MD2DOCX_OBJ = ${MD2DOCX_SRC:.c=.o}

DOCX2MD_SRC = docx2md.c util.c miniz.c
DOCX2MD_OBJ = ${DOCX2MD_SRC:.c=.o}

all: options doctxt md2docx docx2md

options:
	@echo doctxt build options:
	@echo "CFLAGS	= ${CFLAGS}"
	@echo "LDFLAGS	= ${LDFLAGS}"
	@echo "CC		= ${CC}"

${OBJ}: config.mk

.c.o:
	@echo CC $<
	@${CC} -c ${CFLAGS} $<

doctxt: ${OBJ}
	@echo CC -o $@
	@${CC} -o $@ ${OBJ} ${LDFLAGS}

md2docx: md2docx.o util.o miniz.o md4c.o
	@echo CC -o $@
	@${CC} -o $@ md2docx.o util.o miniz.o md4c.o ${LDFLAGS}

docx2md: docx2md.o util.o miniz.o
	@echo CC -o $@
	@${CC} -o $@ docx2md.o util.o miniz.o ${LDFLAGS}

clean:
	@echo cleaning
	@rm -f doctxt md2docx docx2md ${OBJ} md2docx.o docx2md.o doctxt-${VERSION}.tar.gz

dist: clean
	@echo creating dist tarball
	@mkdir -p doctxt-${VERSION}
	@cp -R LICENSE Makefile config.mk ${SRC} doctxt-${VERSION}
	@tar -cf doctxt-${VERSION}.tar doctxt-${VERSION}
	@gzip doctxt-${VERSION}.tar
	@rm -rf doctxt-${VERSION}

install: all
	@echo installing executable files to ${DESTDIR}${PREFIX}/bin
	@mkdir -p ${DESTDIR}${PREFIX}/bin
	@cp -f doctxt ${DESTDIR}${PREFIX}/bin
	@chmod 755 ${DESTDIR}${PREFIX}/bin/doctxt
	@cp -f md2docx ${DESTDIR}${PREFIX}/bin
	@chmod 755 ${DESTDIR}${PREFIX}/bin/md2docx
	@cp -f docx2md ${DESTDIR}${PREFIX}/bin
	@chmod 755 ${DESTDIR}${PREFIX}/bin/docx2md

uninstall:
	@echo removing executable files from ${DESTDIR}${PREFIX}/bin
	@rm -f ${DESTDIR}${PREFIX}/bin/doctxt
	@rm -f ${DESTDIR}${PREFIX}/bin/md2docx
	@rm -f ${DESTDIR}${PREFIX}/bin/docx2md


.PHONY: all options clean install uninstall
