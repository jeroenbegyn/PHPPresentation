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
 * @link        https://github.com/PHPOffice/PHPPresentation
 * @copyright   2009-2015 PHPPresentation contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

declare(strict_types=1);

namespace PhpOffice\PhpPresentation\Shape;


/**
 * Video element
 */
class Video extends Media
{
	
	
	public function __construct() {
		parent::__construct();
		
		// settings
		$this->setName('Video')
		->setDescription('Video')
		->setResizeProportional(false);
	}
	
	
	/**
	 * @return string
	 */
	public function getMimeType(): string
	{
		switch (strtolower($this->getExtension())) {
			case 'mp4':
				$mimetype = 'video/mp4';
				break;
			case 'ogv':
				$mimetype = 'video/ogg';
				break;
			case 'wmv':
				$mimetype = 'video/x-ms-wmv';
				break;
			default:
				$mimetype = 'application/octet-stream';
		}
		return $mimetype;
	}
	
}
