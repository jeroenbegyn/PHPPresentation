<?php

/**
 * This file is part of PHPPresentation - A pure PHP library for reading and writing
 * presentations documents.
 *
 * PHPPresentation is free software distributed under the terms of the GNU Lesser
 * General Public License version 3 as published by the Free Software Foundation.
 *
 * For the full copyright and license information, please read the LICENSE
 * file that was distributed with this source code. For the full list of
 * contributors, visit https://github.com/PHPOffice/PHPPresentation/contributors.
 *
 * @see        https://github.com/PHPOffice/PHPPresentation
 *
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

declare(strict_types=1);

namespace PhpOffice\PhpPresentation\Writer\PowerPoint2007;

use PhpOffice\Common\Adapter\Zip\ZipInterface;
use PhpOffice\Common\Drawing as CommonDrawing;
use PhpOffice\Common\XMLWriter;
use PhpOffice\PhpPresentation\Shape\Audio;
use PhpOffice\PhpPresentation\Shape\Video;
use PhpOffice\PhpPresentation\Shape\Chart as ShapeChart;
use PhpOffice\PhpPresentation\Shape\Comment;
use PhpOffice\PhpPresentation\Shape\Drawing as ShapeDrawing;
use PhpOffice\PhpPresentation\Shape\Group;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Shape\RichText\Run;
use PhpOffice\PhpPresentation\Shape\RichText\TextElement;
use PhpOffice\PhpPresentation\Shape\Table as ShapeTable;
use PhpOffice\PhpPresentation\ShapeContainerInterface;
use PhpOffice\PhpPresentation\Slide;
use PhpOffice\PhpPresentation\Slide\Background\Image;
use PhpOffice\PhpPresentation\Slide\Note;

class PptSlides extends AbstractSlide
{
    /**
     * Add slides (drawings, ...) and slide relationships (drawings, ...).
     */
    public function render(): ZipInterface
    {
        foreach ($this->oPresentation->getAllSlides() as $idx => $oSlide) {
            $this->oZip->addFromString('ppt/slides/_rels/slide' . ($idx + 1) . '.xml.rels', $this->writeSlideRelationships($oSlide));
            $this->oZip->addFromString('ppt/slides/slide' . ($idx + 1) . '.xml', $this->writeSlide($oSlide));

            // Add note slide
            if ($oSlide->getNote() instanceof Note) {
                if (count($oSlide->getNote()->getShapeCollection()) > 0) {
                    $this->oZip->addFromString('ppt/notesSlides/notesSlide' . ($idx + 1) . '.xml', $this->writeNote($oSlide->getNote()));
                }
            }

            // Add background image slide
            $oBkgImage = $oSlide->getBackground();
            if ($oBkgImage instanceof Image) {
                $this->oZip->addFromString('ppt/media/' . $oBkgImage->getIndexedFilename((string) $idx), file_get_contents($oBkgImage->getPath()));
            }
        }

        return $this->oZip;
    }

    /**
     * Write slide relationships to XML format.
     *
     * @return string XML Output
     */
    protected function writeSlideRelationships(Slide $pSlide): string
    {
        //@todo Group all getShapeCollection()->getIterator

        // Create XML writer
        $objWriter = new XMLWriter(XMLWriter::STORAGE_MEMORY);

        // XML header
        $objWriter->startDocument('1.0', 'UTF-8', 'yes');

        // Relationships
        $objWriter->startElement('Relationships');
        $objWriter->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');

        // Starting relation id
        $relId = 1;
        $idxSlide = $pSlide->getParent()->getIndex($pSlide);

        // Write drawing relationships FIRST (before slideLayout) so video gets correct rIds
        if (count($pSlide->getShapeCollection()) > 0) {
            $collections = [$pSlide->getShapeCollection()];

            // Loop trough images and write relationships
            while (count($collections)) {
                $collection = array_shift($collections);

                foreach ($collection as $currentShape) {
                    if ($currentShape instanceof Video) {
                        // For video: Write media embed FIRST (rId1), then video file (rId2), then thumbnail (rId3)
                        // This matches PowerPoint's order and our XML references
                        $currentShape->relationId = 'rId' . ($relId + 1); // Video file will be rId+1
                        $this->writeRelationship($objWriter, $relId, 'http://schemas.microsoft.com/office/2007/relationships/media', '../media/' . $currentShape->getIndexedFilename());
                        ++$relId;
                        $this->writeRelationship($objWriter, $relId, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/video', '../media/' . $currentShape->getIndexedFilename());
                        ++$relId;
                        $filename = str_replace('.', '_', $currentShape->getIndexedFilename()) . '_bg.png';
                        $this->writeRelationship($objWriter, $relId, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image', '../media/' . $filename);
                        ++$relId;
                    } elseif ($currentShape instanceof Audio) {
                        // Write relationship for image drawing
                        $currentShape->relationId = 'rId' . $relId;
                        $this->writeRelationship ( $objWriter, $relId, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio', '../media/' . $currentShape->getIndexedFilename () );
                        ++ $relId;
                        $this->writeRelationship ( $objWriter, $relId, 'http://schemas.microsoft.com/office/2007/relationships/media', '../media/' . $currentShape->getIndexedFilename () );
                        ++ $relId;
                        $filename = str_replace ( '.', '_', $currentShape->getIndexedFilename () ) . '_bg.png';
                        $this->writeRelationship ( $objWriter, $relId, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image', '../media/' . $filename );
                        ++ $relId;
                    }
                    elseif ($currentShape instanceof ShapeDrawing\AbstractDrawingAdapter) {
                        // Write relationship for image drawing
                        $this->writeRelationship($objWriter, $relId, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image', '../media/' . $currentShape->getIndexedFilename());
                        $currentShape->relationId = 'rId' . $relId;
                        ++$relId;
                    } elseif ($currentShape instanceof ShapeChart) {
                        // Write relationship for chart drawing
                        $this->writeRelationship($objWriter, $relId, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart', '../charts/' . $currentShape->getIndexedFilename());
                        $currentShape->relationId = 'rId' . $relId;
                        ++$relId;
                    } elseif ($currentShape instanceof ShapeContainerInterface) {
                        $collections[] = $currentShape->getShapeCollection();
                    }
                }
            }
        }

        // Write background relationships?
        $oBackground = $pSlide->getBackground();
        if ($oBackground instanceof Image) {
            $this->writeRelationship($objWriter, $relId, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image', '../media/' . $oBackground->getIndexedFilename((string) $idxSlide));
            $oBackground->relationId = 'rId' . $relId;
            ++$relId;
        }

        // Write hyperlink relationships?
        if (count($pSlide->getShapeCollection()) > 0) {
            // Loop trough hyperlinks and write relationships
            foreach ($pSlide->getShapeCollection() as $shape) {
                // Hyperlink on shape
                if ($shape->hasHyperlink()) {
                    // Write relationship for hyperlink
                    $hyperlink = $shape->getHyperlink();
                    $hyperlink->relationId = 'rId' . $relId;

                    if (!$hyperlink->isInternal()) {
                        $this->writeRelationship($objWriter, $relId, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', $hyperlink->getUrl(), 'External');
                    } else {
                        $this->writeRelationship($objWriter, $relId, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide', 'slide' . $hyperlink->getSlideNumber() . '.xml');
                    }

                    ++$relId;
                }

                // Hyperlink on rich text run
                if ($shape instanceof RichText) {
                    foreach ($shape->getParagraphs() as $paragraph) {
                        foreach ($paragraph->getRichTextElements() as $element) {
                            if ($element instanceof Run || $element instanceof TextElement) {
                                if ($element->hasHyperlink()) {
                                    // Write relationship for hyperlink
                                    $hyperlink = $element->getHyperlink();
                                    $hyperlink->relationId = 'rId' . $relId;

                                    if (!$hyperlink->isInternal()) {
                                        $this->writeRelationship($objWriter, $relId, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', $hyperlink->getUrl(), 'External');
                                    } else {
                                        $this->writeRelationship($objWriter, $relId, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide', 'slide' . $hyperlink->getSlideNumber() . '.xml');
                                    }

                                    ++$relId;
                                }
                            }
                        }
                    }
                }

                // Hyperlink in table
                if ($shape instanceof ShapeTable) {
                    // Rows
                    $countRows = count($shape->getRows());
                    for ($row = 0; $row < $countRows; ++$row) {
                        // Cells in rows
                        $countCells = count($shape->getRow($row)->getCells());
                        for ($cell = 0; $cell < $countCells; ++$cell) {
                            $currentCell = $shape->getRow($row)->getCell($cell);
                            // Paragraphs in cell
                            foreach ($currentCell->getParagraphs() as $paragraph) {
                                // RichText in paragraph
                                foreach ($paragraph->getRichTextElements() as $element) {
                                    // Run or Text in RichText
                                    if ($element instanceof Run || $element instanceof TextElement) {
                                        if ($element->hasHyperlink()) {
                                            // Write relationship for hyperlink
                                            $hyperlink = $element->getHyperlink();
                                            $hyperlink->relationId = 'rId' . $relId;

                                            if (!$hyperlink->isInternal()) {
                                                $this->writeRelationship($objWriter, $relId, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', $hyperlink->getUrl(), 'External');
                                            } else {
                                                $this->writeRelationship($objWriter, $relId, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide', 'slide' . $hyperlink->getSlideNumber() . '.xml');
                                            }

                                            ++$relId;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                if ($shape instanceof Group) {
                    foreach ($shape->getShapeCollection() as $subShape) {
                        // Hyperlink on shape
                        if ($subShape->hasHyperlink()) {
                            // Write relationship for hyperlink
                            $hyperlink = $subShape->getHyperlink();
                            $hyperlink->relationId = 'rId' . $relId;

                            if (!$hyperlink->isInternal()) {
                                $this->writeRelationship($objWriter, $relId, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', $hyperlink->getUrl(), 'External');
                            } else {
                                $this->writeRelationship($objWriter, $relId, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide', 'slide' . $hyperlink->getSlideNumber() . '.xml');
                            }

                            ++$relId;
                        }

                        // Hyperlink on rich text run
                        if ($subShape instanceof RichText) {
                            foreach ($subShape->getParagraphs() as $paragraph) {
                                foreach ($paragraph->getRichTextElements() as $element) {
                                    if ($element instanceof Run || $element instanceof TextElement) {
                                        if ($element->hasHyperlink()) {
                                            // Write relationship for hyperlink
                                            $hyperlink = $element->getHyperlink();
                                            $hyperlink->relationId = 'rId' . $relId;

                                            if (!$hyperlink->isInternal()) {
                                                $this->writeRelationship($objWriter, $relId, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', $hyperlink->getUrl(), 'External');
                                            } else {
                                                $this->writeRelationship($objWriter, $relId, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide', 'slide' . $hyperlink->getSlideNumber() . '.xml');
                                            }

                                            ++$relId;
                                        }
                                    }
                                }
                            }
                        }

                        // Hyperlink in table
                        if ($subShape instanceof ShapeTable) {
                            // Rows
                            $countRows = count($subShape->getRows());
                            for ($row = 0; $row < $countRows; ++$row) {
                                // Cells in rows
                                $countCells = count($subShape->getRow($row)->getCells());
                                for ($cell = 0; $cell < $countCells; ++$cell) {
                                    $currentCell = $subShape->getRow($row)->getCell($cell);
                                    // Paragraphs in cell
                                    foreach ($currentCell->getParagraphs() as $paragraph) {
                                        // RichText in paragraph
                                        foreach ($paragraph->getRichTextElements() as $element) {
                                            // Run or Text in RichText
                                            if ($element instanceof Run || $element instanceof TextElement) {
                                                if ($element->hasHyperlink()) {
                                                    // Write relationship for hyperlink
                                                    $hyperlink = $element->getHyperlink();
                                                    $hyperlink->relationId = 'rId' . $relId;

                                                    if (!$hyperlink->isInternal()) {
                                                        $this->writeRelationship($objWriter, $relId, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', $hyperlink->getUrl(), 'External');
                                                    } else {
                                                        $this->writeRelationship($objWriter, $relId, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide', 'slide' . $hyperlink->getSlideNumber() . '.xml');
                                                    }

                                                    ++$relId;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        // Write comment relationships
        if (count($pSlide->getShapeCollection()) > 0) {
            $hasSlideComment = false;

            // Loop trough images and write relationships
            foreach ($pSlide->getShapeCollection() as $shape) {
                if ($shape instanceof Comment) {
                    $hasSlideComment = true;

                    break;
                } elseif ($shape instanceof Group) {
                    foreach ($shape->getShapeCollection() as $subShape) {
                        if ($subShape instanceof Comment) {
                            $hasSlideComment = true;

                            break 2;
                        }
                    }
                }
            }

            if ($hasSlideComment) {
                $this->writeRelationship($objWriter, $relId, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments', '../comments/comment' . ($idxSlide + 1) . '.xml');
                ++$relId;
            }
        }

        if (count($pSlide->getNote()->getShapeCollection()) > 0) {
            $this->writeRelationship($objWriter, $relId, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide', '../notesSlides/notesSlide' . ($idxSlide + 1) . '.xml');
            ++$relId;
        }

        // Write slideLayout relationship LAST (after all media relationships)
        $layoutId = 1;
        if ($pSlide->getSlideLayout()) {
            $layoutId = $pSlide->getSlideLayout()->layoutNr;
        }
        $this->writeRelationship($objWriter, $relId, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout', '../slideLayouts/slideLayout' . $layoutId . '.xml');

        $objWriter->endElement();

        // Return
        return $objWriter->getData();
    }

    /**
     * Write slide to XML format.
     *
     * @return string XML Output
     */
    protected function writeSlide(Slide $pSlide): string
    {
        // Create XML writer
        $objWriter = new XMLWriter(XMLWriter::STORAGE_MEMORY);

        // XML header
        $objWriter->startDocument('1.0', 'UTF-8', 'yes');

        // p:sld
        $objWriter->startElement('p:sld');
        $objWriter->writeAttribute('xmlns:a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
        $objWriter->writeAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
        $objWriter->writeAttribute('xmlns:p', 'http://schemas.openxmlformats.org/presentationml/2006/main');
        $objWriter->writeAttributeIf(!$pSlide->isVisible(), 'show', 0);

        // p:sld/p:cSld
        $objWriter->startElement('p:cSld');

        // Background
        if ($pSlide->getBackground() instanceof Slide\AbstractBackground) {
            $oBackground = $pSlide->getBackground();
            // p:bg
            $objWriter->startElement('p:bg');

            // p:bgPr
            $objWriter->startElement('p:bgPr');

            if ($oBackground instanceof Slide\Background\Color) {
                // a:solidFill
                $objWriter->startElement('a:solidFill');

                $this->writeColor($objWriter, $oBackground->getColor());

                // > a:solidFill
                $objWriter->endElement();
            }

            if ($oBackground instanceof Image) {
                // a:blipFill
                $objWriter->startElement('a:blipFill');

                // a:blip
                $objWriter->startElement('a:blip');
                $objWriter->writeAttribute('r:embed', $oBackground->relationId);

                // > a:blipFill
                $objWriter->endElement();

                // a:stretch
                $objWriter->startElement('a:stretch');

                // a:fillRect
                $objWriter->writeElement('a:fillRect');

                // > a:stretch
                $objWriter->endElement();

                // > a:blipFill
                $objWriter->endElement();
            }

            // > p:bgPr
            $objWriter->endElement();

            // > p:bg
            $objWriter->endElement();
        }

        // p:spTree
        $objWriter->startElement('p:spTree');

        // p:nvGrpSpPr
        $objWriter->startElement('p:nvGrpSpPr');

        // p:cNvPr
        $objWriter->startElement('p:cNvPr');
        $objWriter->writeAttribute('id', '1');
        $objWriter->writeAttribute('name', '');
        $objWriter->endElement();

        // p:cNvGrpSpPr
        $objWriter->writeElement('p:cNvGrpSpPr', null);

        // p:nvPr
        $objWriter->writeElement('p:nvPr', null);

        $objWriter->endElement();

        // p:grpSpPr
        $objWriter->startElement('p:grpSpPr');

        // a:xfrm
        $objWriter->startElement('a:xfrm');

        // a:off
        $objWriter->startElement('a:off');
        $objWriter->writeAttribute('x', CommonDrawing::pixelsToEmu($pSlide->getOffsetX()));
        $objWriter->writeAttribute('y', CommonDrawing::pixelsToEmu($pSlide->getOffsetY()));
        $objWriter->endElement(); // a:off

        // a:ext
        $objWriter->startElement('a:ext');
        $objWriter->writeAttribute('cx', CommonDrawing::pixelsToEmu($pSlide->getExtentX()));
        $objWriter->writeAttribute('cy', CommonDrawing::pixelsToEmu($pSlide->getExtentY()));
        $objWriter->endElement(); // a:ext

        // a:chOff
        $objWriter->startElement('a:chOff');
        $objWriter->writeAttribute('x', CommonDrawing::pixelsToEmu($pSlide->getOffsetX()));
        $objWriter->writeAttribute('y', CommonDrawing::pixelsToEmu($pSlide->getOffsetY()));
        $objWriter->endElement(); // a:chOff

        // a:chExt
        $objWriter->startElement('a:chExt');
        $objWriter->writeAttribute('cx', CommonDrawing::pixelsToEmu($pSlide->getExtentX()));
        $objWriter->writeAttribute('cy', CommonDrawing::pixelsToEmu($pSlide->getExtentY()));
        $objWriter->endElement(); // a:chExt

        $objWriter->endElement();

        $objWriter->endElement();

        // Loop shapes
        $this->writeShapeCollection($objWriter, $pSlide->getShapeCollection());

        // TODO
        $objWriter->endElement();

        $objWriter->endElement();

        // p:clrMapOvr
        $objWriter->startElement('p:clrMapOvr');
        // p:clrMapOvr\a:masterClrMapping
        $objWriter->writeElement('a:masterClrMapping', null);
        // ##p:clrMapOvr
        $objWriter->endElement();

        $this->writeSlideTransition($objWriter, $pSlide->getTransition());

        $this->writeSlideAnimations($objWriter, $pSlide);

        $objWriter->endElement();

        // Return
        return $objWriter->getData();
    }

    protected function writeSlideAnimations(XMLWriter $objWriter, Slide $oSlide): void
    {
        $arrayAnimations = $oSlide->getAnimations();
        
        // Check if there are any audio or video shapes in the slide
        $audioShapes = [];
        $videoShapes = [];
        $shapeId = 1;
        foreach ($oSlide->getShapeCollection() as $shape) {
            ++$shapeId;
            if ($shape instanceof Audio) {
                $audioShapes[] = ['shape' => $shape, 'id' => $shapeId];
            } elseif ($shape instanceof Video) {
                $videoShapes[] = ['shape' => $shape, 'id' => $shapeId];
            }
        }
        
        // If no animations and no media, return early
        if (empty($arrayAnimations) && empty($audioShapes) && empty($videoShapes)) {
            return;
        }
        
        // If we have media but no animations, write media timing
        if (empty($arrayAnimations) && (!empty($audioShapes) || !empty($videoShapes))) {
            $this->writeMediaTiming($objWriter, $audioShapes, $videoShapes);
            return;
        }

        // Variables
        $shapeId = 1;
        $idCount = 1;
        $hashToIdMap = [];
        $arrayAnimationIds = [];

        foreach ($oSlide->getShapeCollection() as $shape) {
            $hashToIdMap[$shape->getHashCode()] = ++$shapeId;
        }
        foreach ($arrayAnimations as $oAnimation) {
            foreach ($oAnimation->getShapeCollection() as $oShape) {
                $arrayAnimationIds[] = $hashToIdMap[$oShape->getHashCode()];
            }
        }

        // p:timing
        $objWriter->startElement('p:timing');
        // p:timing/p:tnLst
        $objWriter->startElement('p:tnLst');
        // p:timing/p:tnLst/p:par
        $objWriter->startElement('p:par');
        // p:timing/p:tnLst/p:par/p:cTn
        $objWriter->startElement('p:cTn');
        $objWriter->writeAttribute('id', $idCount++);
        $objWriter->writeAttribute('dur', 'indefinite');
        $objWriter->writeAttribute('restart', 'never');
        $objWriter->writeAttribute('nodeType', 'tmRoot');
        // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst
        $objWriter->startElement('p:childTnLst');
        // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq
        $objWriter->startElement('p:seq');
        $objWriter->writeAttribute('concurrent', '1');
        $objWriter->writeAttribute('nextAc', 'seek');
        // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:cTn
        $objWriter->startElement('p:cTn');
        $objWriter->writeAttribute('id', $idCount++);
        $objWriter->writeAttribute('dur', 'indefinite');
        $objWriter->writeAttribute('nodeType', 'mainSeq');
        // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:cTn/p:childTnLst
        $objWriter->startElement('p:childTnLst');

        // Each animation has multiple shapes
        foreach ($arrayAnimations as $oAnimation) {
            // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:cTn/p:childTnLst/p:par
            $objWriter->startElement('p:par');
            // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:cTn/p:childTnLst/p:par/p:cTn
            $objWriter->startElement('p:cTn');
            $objWriter->writeAttribute('id', $idCount++);
            $objWriter->writeAttribute('fill', 'hold');
            // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:cTn/p:childTnLst/p:par/p:cTn/p:stCondLst
            $objWriter->startElement('p:stCondLst');
            // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:cTn/p:childTnLst/p:par/p:cTn/p:stCondLst/p:cond
            $objWriter->startElement('p:cond');
            $objWriter->writeAttribute('delay', 'indefinite');
            $objWriter->endElement();
            // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:cTn/p:childTnLst/p:par/p:cTn\##p:stCondLst
            $objWriter->endElement();
            // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:cTn/p:childTnLst/p:par/p:cTn/p:childTnLst
            $objWriter->startElement('p:childTnLst');
            // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:cTn/p:childTnLst/p:par/p:cTn/p:childTnLst/p:par
            $objWriter->startElement('p:par');
            // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:cTn/p:childTnLst/p:par/p:cTn/p:childTnLst/p:par/p:cTn
            $objWriter->startElement('p:cTn');
            $objWriter->writeAttribute('id', $idCount++);
            $objWriter->writeAttribute('fill', 'hold');
            // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:cTn/p:childTnLst/p:par/p:cTn/p:childTnLst/p:par/p:cTn/p:stCondLst
            $objWriter->startElement('p:stCondLst');
            // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:cTn/p:childTnLst/p:par/p:cTn/p:childTnLst/p:par/p:cTn/p:stCondLst/p:cond
            $objWriter->startElement('p:cond');
            $objWriter->writeAttribute('delay', '0');
            $objWriter->endElement();
            // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:cTn/p:childTnLst/p:par/p:cTn/p:childTnLst/p:par/p:cTn\##p:stCondLst
            $objWriter->endElement();
            // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:cTn/p:childTnLst/p:par/p:cTn/p:childTnLst/p:par/p:cTn/p:childTnLst
            $objWriter->startElement('p:childTnLst');

            $firstAnimation = true;
            foreach ($oAnimation->getShapeCollection() as $oShape) {
                $nodeType = $firstAnimation ? 'clickEffect' : 'withEffect';
                $shapeId = $hashToIdMap[$oShape->getHashCode()];

                // p:par
                $objWriter->startElement('p:par');
                // p:par/p:cTn
                $objWriter->startElement('p:cTn');
                $objWriter->writeAttribute('id', $idCount++);
                $objWriter->writeAttribute('presetID', '1');
                $objWriter->writeAttribute('presetClass', 'entr');
                $objWriter->writeAttribute('fill', 'hold');
                $objWriter->writeAttribute('presetSubtype', '0');
                $objWriter->writeAttribute('grpId', '0');
                $objWriter->writeAttribute('nodeType', $nodeType);
                // p:par/p:cTn/p:stCondLst
                $objWriter->startElement('p:stCondLst');
                // p:par/p:cTn/p:stCondLst/p:cond
                $objWriter->startElement('p:cond');
                $objWriter->writeAttribute('delay', '0');
                $objWriter->endElement();
                // p:par/p:cTn\##p:stCondLst
                $objWriter->endElement();
                // p:par/p:cTn/p:childTnLst
                $objWriter->startElement('p:childTnLst');
                // p:par/p:cTn/p:childTnLst/p:set
                $objWriter->startElement('p:set');
                // p:par/p:cTn/p:childTnLst/p:set/p:cBhvr
                $objWriter->startElement('p:cBhvr');
                // p:par/p:cTn/p:childTnLst/p:set/p:cBhvr/p:cTn
                $objWriter->startElement('p:cTn');
                $objWriter->writeAttribute('id', $idCount++);
                $objWriter->writeAttribute('dur', '1');
                $objWriter->writeAttribute('fill', 'hold');
                // p:par/p:cTn/p:childTnLst/p:set/p:cBhvr/p:cTn/p:stCondLst
                $objWriter->startElement('p:stCondLst');
                // p:par/p:cTn/p:childTnLst/p:set/p:cBhvr/p:cTn/p:stCondLst/p:cond
                $objWriter->startElement('p:cond');
                $objWriter->writeAttribute('delay', '0');
                $objWriter->endElement();
                // p:par/p:cTn/p:childTnLst/p:set/p:cBhvr/p:cTn\##p:stCondLst
                $objWriter->endElement();
                // p:par/p:cTn/p:childTnLst/p:set/p:cBhvr\##p:cTn
                $objWriter->endElement();
                // p:par/p:cTn/p:childTnLst/p:set/p:cBhvr/p:tgtEl
                $objWriter->startElement('p:tgtEl');
                // p:par/p:cTn/p:childTnLst/p:set/p:cBhvr/p:tgtEl/p:spTgt
                $objWriter->startElement('p:spTgt');
                $objWriter->writeAttribute('spid', $shapeId);
                $objWriter->endElement();
                // p:par/p:cTn/p:childTnLst/p:set/p:cBhvr\##p:tgtEl
                $objWriter->endElement();
                // p:par/p:cTn/p:childTnLst/p:set/p:cBhvr/p:attrNameLst
                $objWriter->startElement('p:attrNameLst');
                // p:par/p:cTn/p:childTnLst/p:set/p:cBhvr/p:attrNameLst/p:attrName
                $objWriter->writeElement('p:attrName', 'style.visibility');
                // p:par/p:cTn/p:childTnLst/p:set/p:cBhvr\##p:attrNameLst
                $objWriter->endElement();
                // p:par/p:cTn/p:childTnLst/p:set\##p:cBhvr
                $objWriter->endElement();
                // p:par/p:cTn/p:childTnLst/p:set/p:to
                $objWriter->startElement('p:to');
                // p:par/p:cTn/p:childTnLst/p:set/p:to/p:strVal
                $objWriter->startElement('p:strVal');
                $objWriter->writeAttribute('val', 'visible');
                $objWriter->endElement();
                // p:par/p:cTn/p:childTnLst/p:set\##p:to
                $objWriter->endElement();
                // p:par/p:cTn/p:childTnLst\##p:set
                $objWriter->endElement();
                // p:par/p:cTn\##p:childTnLst
                $objWriter->endElement();
                // p:par\##p:cTn
                $objWriter->endElement();
                // ##p:par
                $objWriter->endElement();

                $firstAnimation = false;
            }

            // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:cTn/p:childTnLst/p:par/p:cTn/p:childTnLst/p:par/p:cTn\##p:childTnLst
            $objWriter->endElement();
            // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:cTn/p:childTnLst/p:par/p:cTn/p:childTnLst/p:par\##p:cTn
            $objWriter->endElement();
            // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:cTn/p:childTnLst/p:par/p:cTn/p:childTnLst\##p:par
            $objWriter->endElement();
            // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:cTn/p:childTnLst/p:par/p:cTn\##p:childTnLst
            $objWriter->endElement();
            // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:cTn/p:childTnLst/p:par\##p:cTn
            $objWriter->endElement();
            // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:cTn/p:childTnLst\##p:par
            $objWriter->endElement();
        }

        // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:cTn\##p:childTnLst
        $objWriter->endElement();
        // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq\##p:cTn
        $objWriter->endElement();
        // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:prevCondLst
        $objWriter->startElement('p:prevCondLst');
        // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:prevCondLst/p:cond
        $objWriter->startElement('p:cond');
        $objWriter->writeAttribute('evt', 'onPrev');
        $objWriter->writeAttribute('delay', '0');
        // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:prevCondLst/p:cond/p:tgtEl
        $objWriter->startElement('p:tgtEl');
        // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:prevCondLst/p:cond/p:tgtEl/p:sldTgt
        $objWriter->writeElement('p:sldTgt', null);
        // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:prevCondLst/p:cond\##p:tgtEl
        $objWriter->endElement();
        // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:prevCondLst\##p:cond
        $objWriter->endElement();
        // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq\##p:prevCondLst
        $objWriter->endElement();
        // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:nextCondLst
        $objWriter->startElement('p:nextCondLst');
        // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:nextCondLst/p:cond
        $objWriter->startElement('p:cond');
        $objWriter->writeAttribute('evt', 'onNext');
        $objWriter->writeAttribute('delay', '0');
        // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:nextCondLst/p:cond/p:tgtEl
        $objWriter->startElement('p:tgtEl');
        // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:nextCondLst/p:cond/p:tgtEl/p:sldTgt
        $objWriter->writeElement('p:sldTgt', null);
        // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:nextCondLst/p:cond\##p:tgtEl
        $objWriter->endElement();
        // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:nextCondLst\##p:cond
        $objWriter->endElement();
        // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq\##p:nextCondLst
        $objWriter->endElement();
        // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst\##p:seq
        $objWriter->endElement();
        // p:timing/p:tnLst/p:par/p:cTn\##p:childTnLst
        $objWriter->endElement();
        // p:timing/p:tnLst/p:par\##p:cTn
        $objWriter->endElement();
        // p:timing/p:tnLst\##p:par
        $objWriter->endElement();
        // p:timing\##p:tnLst
        $objWriter->endElement();

        // p:timing/p:bldLst
        $objWriter->startElement('p:bldLst');

        // Add in ids of all shapes in this slides animations
        foreach ($arrayAnimationIds as $id) {
            // p:timing/p:bldLst/p:bldP
            $objWriter->startElement('p:bldP');
            $objWriter->writeAttribute('spid', $id);
            $objWriter->writeAttribute('grpId', 0);
            $objWriter->endELement();
        }

        // p:timing\##p:bldLst
        $objWriter->endElement();

        // ##p:timing
        $objWriter->endElement();
    }

    /**
     * Write media timing section for slides with audio/video but no animations
     *
     * @param XMLWriter $objWriter
     * @param array $audioShapes Array of audio shapes with their IDs
     * @param array $videoShapes Array of video shapes with their IDs
     */
    protected function writeMediaTiming(XMLWriter $objWriter, array $audioShapes, array $videoShapes): void
    {
        // Combine all media shapes
        $mediaShapes = array_merge($audioShapes, $videoShapes);
        // p:timing
        $objWriter->startElement('p:timing');
        // p:timing/p:tnLst
        $objWriter->startElement('p:tnLst');
        // p:timing/p:tnLst/p:par
        $objWriter->startElement('p:par');
        // p:timing/p:tnLst/p:par/p:cTn
        $objWriter->startElement('p:cTn');
        $objWriter->writeAttribute('id', '1');
        $objWriter->writeAttribute('dur', 'indefinite');
        $objWriter->writeAttribute('restart', 'never');
        $objWriter->writeAttribute('nodeType', 'tmRoot');
        // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst
        $objWriter->startElement('p:childTnLst');
        
        // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq
        $objWriter->startElement('p:seq');
        $objWriter->writeAttribute('concurrent', '1');
        $objWriter->writeAttribute('nextAc', 'seek');
        // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:cTn
        $objWriter->startElement('p:cTn');
        $objWriter->writeAttribute('id', '2');
        $objWriter->writeAttribute('dur', 'indefinite');
        $objWriter->writeAttribute('nodeType', 'mainSeq');
        // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:cTn/p:childTnLst
        $objWriter->startElement('p:childTnLst');
        
        $idCount = 3;
        foreach ($mediaShapes as $mediaInfo) {
            $mediaShape = $mediaInfo['shape'];
            $shapeId = $mediaInfo['id'];
            $isVideo = $mediaShape instanceof Video;
            
            // p:par
            $objWriter->startElement('p:par');
            // p:par/p:cTn
            $objWriter->startElement('p:cTn');
            $objWriter->writeAttribute('id', $idCount++);
            $objWriter->writeAttribute('fill', 'hold');
            // p:par/p:cTn/p:stCondLst
            $objWriter->startElement('p:stCondLst');
            // p:par/p:cTn/p:stCondLst/p:cond
            $objWriter->startElement('p:cond');
            $objWriter->writeAttribute('delay', 'indefinite');
            $objWriter->endElement(); // p:cond
            $objWriter->endElement(); // p:stCondLst
            
            // p:par/p:cTn/p:childTnLst
            $objWriter->startElement('p:childTnLst');
            // p:par/p:cTn/p:childTnLst/p:par
            $objWriter->startElement('p:par');
            // p:par/p:cTn/p:childTnLst/p:par/p:cTn
            $objWriter->startElement('p:cTn');
            $objWriter->writeAttribute('id', $idCount++);
            $objWriter->writeAttribute('fill', 'hold');
            // p:par/p:cTn/p:childTnLst/p:par/p:cTn/p:stCondLst
            $objWriter->startElement('p:stCondLst');
            // p:par/p:cTn/p:childTnLst/p:par/p:cTn/p:stCondLst/p:cond
            $objWriter->startElement('p:cond');
            $objWriter->writeAttribute('delay', '0');
            $objWriter->endElement(); // p:cond
            $objWriter->endElement(); // p:stCondLst
            
            // p:par/p:cTn/p:childTnLst/p:par/p:cTn/p:childTnLst
            $objWriter->startElement('p:childTnLst');
            // p:par/p:cTn/p:childTnLst/p:par/p:cTn/p:childTnLst/p:par
            $objWriter->startElement('p:par');
            // p:par/p:cTn/p:childTnLst/p:par/p:cTn/p:childTnLst/p:par/p:cTn
            $objWriter->startElement('p:cTn');
            $objWriter->writeAttribute('id', $idCount++);
            $objWriter->writeAttribute('presetID', '1');
            $objWriter->writeAttribute('presetClass', 'mediacall');
            $objWriter->writeAttribute('presetSubtype', '0');
            $objWriter->writeAttribute('fill', 'hold');
            $objWriter->writeAttribute('nodeType', 'clickEffect');
            // p:par/p:cTn/p:childTnLst/p:par/p:cTn/p:childTnLst/p:par/p:cTn/p:stCondLst
            $objWriter->startElement('p:stCondLst');
            // p:par/p:cTn/p:childTnLst/p:par/p:cTn/p:childTnLst/p:par/p:cTn/p:stCondLst/p:cond
            $objWriter->startElement('p:cond');
            $objWriter->writeAttribute('delay', '0');
            $objWriter->endElement(); // p:cond
            $objWriter->endElement(); // p:stCondLst
            
            // p:par/p:cTn/p:childTnLst/p:par/p:cTn/p:childTnLst/p:par/p:cTn/p:childTnLst
            $objWriter->startElement('p:childTnLst');
            // p:cmd
            $objWriter->startElement('p:cmd');
            $objWriter->writeAttribute('type', 'call');
            $objWriter->writeAttribute('cmd', 'playFrom(0.0)');
            // p:cmd/p:cBhvr
            $objWriter->startElement('p:cBhvr');
            // p:cmd/p:cBhvr/p:cTn
            $objWriter->startElement('p:cTn');
            $objWriter->writeAttribute('id', $idCount++);
            
            // Get audio duration if available (in milliseconds)
            // For now, we'll use a default duration
            $duration = '25032'; // Default duration, you might want to get this from the audio file
            $objWriter->writeAttribute('dur', $duration);
            $objWriter->writeAttribute('fill', 'hold');
            $objWriter->endElement(); // p:cTn
            
            // p:cmd/p:cBhvr/p:tgtEl
            $objWriter->startElement('p:tgtEl');
            // p:cmd/p:cBhvr/p:tgtEl/p:spTgt
            $objWriter->startElement('p:spTgt');
            $objWriter->writeAttribute('spid', $shapeId);
            $objWriter->endElement(); // p:spTgt
            $objWriter->endElement(); // p:tgtEl
            $objWriter->endElement(); // p:cBhvr
            $objWriter->endElement(); // p:cmd
            
            $objWriter->endElement(); // p:childTnLst
            $objWriter->endElement(); // p:cTn
            $objWriter->endElement(); // p:par (inner)
            $objWriter->endElement(); // p:childTnLst
            $objWriter->endElement(); // p:cTn
            $objWriter->endElement(); // p:par (middle)
            $objWriter->endElement(); // p:childTnLst
            $objWriter->endElement(); // p:cTn
            $objWriter->endElement(); // p:par (outer)
        }
        
        $objWriter->endElement(); // p:childTnLst
        $objWriter->endElement(); // p:cTn
        
        // p:prevCondLst
        $objWriter->startElement('p:prevCondLst');
        $objWriter->startElement('p:cond');
        $objWriter->writeAttribute('evt', 'onPrev');
        $objWriter->writeAttribute('delay', '0');
        $objWriter->startElement('p:tgtEl');
        $objWriter->writeElement('p:sldTgt', null);
        $objWriter->endElement(); // p:tgtEl
        $objWriter->endElement(); // p:cond
        $objWriter->endElement(); // p:prevCondLst
        
        // p:nextCondLst
        $objWriter->startElement('p:nextCondLst');
        $objWriter->startElement('p:cond');
        $objWriter->writeAttribute('evt', 'onNext');
        $objWriter->writeAttribute('delay', '0');
        $objWriter->startElement('p:tgtEl');
        $objWriter->writeElement('p:sldTgt', null);
        $objWriter->endElement(); // p:tgtEl
        $objWriter->endElement(); // p:cond
        $objWriter->endElement(); // p:nextCondLst
        
        $objWriter->endElement(); // p:seq
        
        // Media nodes (p:audio or p:video)
        foreach ($mediaShapes as $mediaInfo) {
            $mediaShape = $mediaInfo['shape'];
            $shapeId = $mediaInfo['id'];
            $isVideo = $mediaShape instanceof Video;
            
            // Write p:audio or p:video
            $objWriter->startElement($isVideo ? 'p:video' : 'p:audio');
            // p:cMediaNode
            $objWriter->startElement('p:cMediaNode');
            $objWriter->writeAttribute('vol', '80000');
            // p:cMediaNode/p:cTn
            $objWriter->startElement('p:cTn');
            $objWriter->writeAttribute('id', $idCount++);
            $objWriter->writeAttribute('fill', 'hold');
            $objWriter->writeAttribute('display', '0');
            
            // p:cMediaNode/p:cTn/p:stCondLst
            $objWriter->startElement('p:stCondLst');
            $objWriter->startElement('p:cond');
            $objWriter->writeAttribute('delay', 'indefinite');
            $objWriter->endElement(); // p:cond
            $objWriter->endElement(); // p:stCondLst
            
            // For video only, add p:endCondLst (audio has it but video doesn't need onStopAudio)
            if (!$isVideo) {
                // p:cMediaNode/p:cTn/p:endCondLst
                $objWriter->startElement('p:endCondLst');
                $objWriter->startElement('p:cond');
                $objWriter->writeAttribute('evt', 'onStopAudio');
                $objWriter->writeAttribute('delay', '0');
                $objWriter->startElement('p:tgtEl');
                $objWriter->writeElement('p:sldTgt', null);
                $objWriter->endElement(); // p:tgtEl
                $objWriter->endElement(); // p:cond
                $objWriter->endElement(); // p:endCondLst
            }
            
            $objWriter->endElement(); // p:cTn
            
            // p:cMediaNode/p:tgtEl
            $objWriter->startElement('p:tgtEl');
            $objWriter->startElement('p:spTgt');
            $objWriter->writeAttribute('spid', $shapeId);
            $objWriter->endElement(); // p:spTgt
            $objWriter->endElement(); // p:tgtEl
            
            $objWriter->endElement(); // p:cMediaNode
            $objWriter->endElement(); // p:audio or p:video
        }
        
        // For video, add interactive sequence for play/pause toggle
        foreach ($videoShapes as $videoInfo) {
            $shapeId = $videoInfo['id'];
            
            // p:seq for interactive controls
            $objWriter->startElement('p:seq');
            $objWriter->writeAttribute('concurrent', '1');
            $objWriter->writeAttribute('nextAc', 'seek');
            // p:seq/p:cTn
            $objWriter->startElement('p:cTn');
            $objWriter->writeAttribute('id', $idCount++);
            $objWriter->writeAttribute('restart', 'whenNotActive');
            $objWriter->writeAttribute('fill', 'hold');
            $objWriter->writeAttribute('evtFilter', 'cancelBubble');
            $objWriter->writeAttribute('nodeType', 'interactiveSeq');
            
            // p:seq/p:cTn/p:stCondLst
            $objWriter->startElement('p:stCondLst');
            $objWriter->startElement('p:cond');
            $objWriter->writeAttribute('evt', 'onClick');
            $objWriter->writeAttribute('delay', '0');
            $objWriter->startElement('p:tgtEl');
            $objWriter->startElement('p:spTgt');
            $objWriter->writeAttribute('spid', $shapeId);
            $objWriter->endElement(); // p:spTgt
            $objWriter->endElement(); // p:tgtEl
            $objWriter->endElement(); // p:cond
            $objWriter->endElement(); // p:stCondLst
            
            // p:seq/p:cTn/p:endSync
            $objWriter->startElement('p:endSync');
            $objWriter->writeAttribute('evt', 'end');
            $objWriter->writeAttribute('delay', '0');
            $objWriter->startElement('p:rtn');
            $objWriter->writeAttribute('val', 'all');
            $objWriter->endElement(); // p:rtn
            $objWriter->endElement(); // p:endSync
            
            // p:seq/p:cTn/p:childTnLst
            $objWriter->startElement('p:childTnLst');
            $objWriter->startElement('p:par');
            $objWriter->startElement('p:cTn');
            $objWriter->writeAttribute('id', $idCount++);
            $objWriter->writeAttribute('fill', 'hold');
            $objWriter->startElement('p:stCondLst');
            $objWriter->startElement('p:cond');
            $objWriter->writeAttribute('delay', '0');
            $objWriter->endElement(); // p:cond
            $objWriter->endElement(); // p:stCondLst
            
            $objWriter->startElement('p:childTnLst');
            $objWriter->startElement('p:par');
            $objWriter->startElement('p:cTn');
            $objWriter->writeAttribute('id', $idCount++);
            $objWriter->writeAttribute('fill', 'hold');
            $objWriter->startElement('p:stCondLst');
            $objWriter->startElement('p:cond');
            $objWriter->writeAttribute('delay', '0');
            $objWriter->endElement(); // p:cond
            $objWriter->endElement(); // p:stCondLst
            
            $objWriter->startElement('p:childTnLst');
            $objWriter->startElement('p:par');
            $objWriter->startElement('p:cTn');
            $objWriter->writeAttribute('id', $idCount++);
            $objWriter->writeAttribute('presetID', '2');
            $objWriter->writeAttribute('presetClass', 'mediacall');
            $objWriter->writeAttribute('presetSubtype', '0');
            $objWriter->writeAttribute('fill', 'hold');
            $objWriter->writeAttribute('nodeType', 'clickEffect');
            $objWriter->startElement('p:stCondLst');
            $objWriter->startElement('p:cond');
            $objWriter->writeAttribute('delay', '0');
            $objWriter->endElement(); // p:cond
            $objWriter->endElement(); // p:stCondLst
            
            $objWriter->startElement('p:childTnLst');
            $objWriter->startElement('p:cmd');
            $objWriter->writeAttribute('type', 'call');
            $objWriter->writeAttribute('cmd', 'togglePause');
            $objWriter->startElement('p:cBhvr');
            $objWriter->startElement('p:cTn');
            $objWriter->writeAttribute('id', $idCount++);
            $objWriter->writeAttribute('dur', '1');
            $objWriter->writeAttribute('fill', 'hold');
            $objWriter->endElement(); // p:cTn
            $objWriter->startElement('p:tgtEl');
            $objWriter->startElement('p:spTgt');
            $objWriter->writeAttribute('spid', $shapeId);
            $objWriter->endElement(); // p:spTgt
            $objWriter->endElement(); // p:tgtEl
            $objWriter->endElement(); // p:cBhvr
            $objWriter->endElement(); // p:cmd
            $objWriter->endElement(); // p:childTnLst
            $objWriter->endElement(); // p:cTn (presetID=2)
            $objWriter->endElement(); // p:par
            $objWriter->endElement(); // p:childTnLst
            $objWriter->endElement(); // p:cTn (id=10)
            $objWriter->endElement(); // p:par
            $objWriter->endElement(); // p:childTnLst
            $objWriter->endElement(); // p:cTn (id=9)
            $objWriter->endElement(); // p:par
            $objWriter->endElement(); // p:childTnLst
            $objWriter->endElement(); // p:cTn (id=9)
            $objWriter->endElement(); // p:par
            $objWriter->endElement(); // p:childTnLst
            
            // p:seq/p:nextCondLst (sibling of p:cTn, not child)
            $objWriter->startElement('p:nextCondLst');
            $objWriter->startElement('p:cond');
            $objWriter->writeAttribute('evt', 'onClick');
            $objWriter->writeAttribute('delay', '0');
            $objWriter->startElement('p:tgtEl');
            $objWriter->startElement('p:spTgt');
            $objWriter->writeAttribute('spid', $shapeId);
            $objWriter->endElement(); // p:spTgt
            $objWriter->endElement(); // p:tgtEl
            $objWriter->endElement(); // p:cond
            $objWriter->endElement(); // p:nextCondLst
            
            $objWriter->endElement(); // p:seq
        }
        
        $objWriter->endElement(); // p:childTnLst
        $objWriter->endElement(); // p:cTn
        $objWriter->endElement(); // p:par
        $objWriter->endElement(); // p:tnLst
        $objWriter->endElement(); // p:timing
    }
}
