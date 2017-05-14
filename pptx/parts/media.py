# encoding: utf-8

"""MediaPart and related objects."""

import hashlib
import os

from ..opc.package import Part
from ..util import lazyproperty


class MediaPart(Part):
    """A media part.

    A media part generally has a partname matching the regex
    ``ppt/media/media[1-9][0-9]*.*``.
    """

    def __init__(self, partname, content_type, blob, package):
        super(MediaPart, self).__init__(
            partname, content_type, blob, package
        )

    @classmethod
    def load(cls, partname, content_type, blob, package):
        return cls(partname, content_type, blob, package)

    @classmethod
    def new(cls, package, media):
        """Return new |MediaPart| instance containing *media*.

        *media* must be a |Media| object.
        """
        partname = package.next_media_partname(media.ext)
        return cls(partname, media.content_type, media.blob, package)

    @lazyproperty
    def sha1(self):
        """
        The SHA1 hash digest for the media binary of this media part, like:
        ``'1be010ea47803b00e140b852765cdf84f491da47'``.
        """
        return hashlib.sha1(self._blob).hexdigest()


class Media(object):
    """Immutable value object representing media such as MP4."""

    def __init__(self, blob, filename, content_type):
        super(Media, self).__init__()
        self._blob = blob
        self._filename = filename
        self._content_type = content_type

    @classmethod
    def from_file(cls, media_file, content_type):
        """Return a new |Media| object loaded from *media_file*.

        *media_file* must be a filesystem path (string). *content_type* is
        the MIME type string, such as 'video/unknown'.
        """
        with open(media_file, 'rb') as f:
            blob = f.read()
        filename = os.path.basename(media_file)
        return cls(blob, filename, content_type)

    @property
    def blob(self):
        """The bytestream of the media "file"."""
        return self._blob

    @property
    def content_type(self):
        """
        MIME-type of this media, e.g. ``'media/jpeg'``.
        """
        return self._content_type

    @property
    def ext(self):
        """Return the file extension for this media e.g. `mp4`."""
        _, dotted_ext = os.path.splitext(self._filename)
        undotted_ext = dotted_ext.lstrip('.')
        return undotted_ext

    @lazyproperty
    def sha1(self):
        """SHA1 hash digest of the media blob."""
        return hashlib.sha1(self._blob).hexdigest()


import base64
speaker_image_bytes = base64.b64decode(
    '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'
    'umLMN+O8AX65uqCMleo4AAAAASUVORK5CYII='
)
