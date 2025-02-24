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
 * @copyright   2009-2015 PHPPresentation contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */
declare ( strict_types = 1 )
	;

namespace PhpOffice\PhpPresentation\Slide;

use PhpOffice\PhpPresentation\AbstractShape;

class Animation {
	/**
	 *
	 * @var array<AbstractShape>
	 */
	protected $shapeCollection = array ();

	/**
	 * Duration of animation
	 *
	 * @var integer
	 */
	private $duration = 30;

	// blinds(horizontal)
	// blinds(vertical)
	// box(in)
	// box(out)
	// checkerboard(across)
	// checkerboard(down)
	// circle(in)
	// circle(out)
	// diamond(in)
	// diamond(out)
	// dissolve
	// fade
	// slide(fromTop)
	// slide(fromBottom)
	// slide(fromLeft)
	// slide(fromRight)
	// plus(in)
	// plus(out)
	// barn(inVertical)
	// barn(inHorizontal)
	// barn(outVertical)
	// barn(outHorizontal)
	// randombar(horizontal)
	// randombar(vertical)
	// strips(downLeft)
	// strips(upLeft)
	// strips(downRight)
	// strips(upRight)
	// wedge
	// wheel(1)
	// wheel(2)
	// wheel(3)
	// wheel(4)
	// wheel(8)
	// wipe(right)
	// wipe(left)
	// wipe(up)
	// wipe(down)
	private $animEffectFilter = 'wipe(down)';

	// in
	// out
	private $animEffectTransition = 'out';

	/**
	 *
	 * @return Animation
	 */
	public function addShape(AbstractShape $shape) {
		$this->shapeCollection [] = $shape;

		return $this;
	}

	/**
	 *
	 * @return array<AbstractShape>
	 */
	public function getShapeCollection(): array {
		return $this->shapeCollection;
	}

	/**
	 *
	 * @param array<AbstractShape> $array
	 *
	 * @return Animation
	 */
	public function setShapeCollection(array $array = array ()) {
		$this->shapeCollection = $array;

		return $this;
	}
	/**
	 *
	 * @return number
	 */
	public function getDuration() {
		return $this->duration;
	}
	/**
	 *
	 * @param integer $duration
	 * @return \PhpOffice\PhpPresentation\Slide\Animation
	 */
	public function setDuration( $duration) {
		$this->duration = $duration;
		return $this;
	}
	/**
	 *
	 * @return string
	 */
	public function getAnimEffectFilter() {
		return $this->animEffectFilter;
	}
	/**
	 *
	 * @param string $animEffectFilter
	 * @return \PhpOffice\PhpPresentation\Slide\Animation
	 */
	public function setAnimEffectFilter($animEffectFilter) {
		$this->animEffectFilter = $animEffectFilter;
		return $this;
	}
	/**
	 *
	 * @return string
	 */
	public function getAnimEffectTransition() {
		return $this->animEffectTransition;
	}
	/**
	 *
	 * @param string $animEffectTransition
	 * @return \PhpOffice\PhpPresentation\Slide\Animation
	 */
	public function setAnimEffectTransition($animEffectTransition) {
		$this->animEffectTransition = $animEffectTransition;
		return $this;
	}

}
