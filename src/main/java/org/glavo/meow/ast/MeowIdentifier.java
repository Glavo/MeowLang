/*
 * Copyright (C) 2025 Glavo
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program.  If not, see <https://www.gnu.org/licenses/>.
 */
package org.glavo.meow.ast;

import org.glavo.meow.Meow;
import org.glavo.meow.MeowContext;
import org.glavo.meow.MeowSymbolMap;
import org.glavo.meow.value.MeowValue;

public record MeowIdentifier(Meow meow) implements MeowExpression {

    @Override
    public MeowValue eval(MeowContext context) {
        return context.getValue(meow);
    }

    @Override
    public String toDebugString() {
        return MeowSymbolMap.INSTANCE.get(meow);
    }
}
